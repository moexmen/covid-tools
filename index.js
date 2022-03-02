'use strict';

let data = {
  inputStats: {
    totalRead: 0,
    invalid: 0,
    duplicate: 0,
  },
  inputUin: new Map(),
  inputColumnHeaders: [],
  stats: {
    retrieved: 0,
    positiveTestResults: 0,
    negativeTestResults: 0,
    pendingTestResults: 0,
    invalidTestResults: 0,
    noTestResults: 0
  },
  progress: 0,
  logs: [],
};

const uinRegex = /^([FGMST])([0-9]{7})([A-Z])$/;
const uinWeights = [2, 7, 6, 5, 4, 3, 2];
const nricChecksum = ['J', 'Z', 'I', 'H', 'G', 'F', 'E', 'D', 'C', 'B', 'A'];
const finFGChecksum = ['X', 'W', 'U', 'T', 'R', 'Q', 'P', 'N', 'M', 'L', 'K'];
const finMChecksum = ['X', 'W', 'U', 'T', 'R', 'Q', 'P', 'N', 'J', 'L', 'K'];
// const baseurlRegex = /^https:\/\/[A-Za-z\.\-]*\/api\/v1$/;
const baseurlRegex = /^https?:\/\/[A-Za-z\.\-:0-9]*\/v2$/;
const keyRegex = /^[A-Za-z0-9\-_\/=+]{8,}$/;

// assume that prefix and digits strings are already checked for length, chars
function uinChecksum(prefix, digits) {
  let sum = 0;
  for (let i = 0; i < 7; i++) {
    sum += digits[i] * uinWeights[i];
  }
  if (prefix == 'T' || prefix == 'G') {
    sum += 4;
  }
  if (prefix == 'M') {
    sum += 3;
  }
  sum = sum % 11;

  let expectedChecksum = '';
  if (prefix == 'S' || prefix == 'T') {
    expectedChecksum = nricChecksum[sum];
  }
  if (prefix == 'F' || prefix == 'G') {
    expectedChecksum = finFGChecksum[sum];
  }
  if (prefix == 'M') {
    expectedChecksum = finMChecksum[sum];
  }

  return expectedChecksum;
}

// this is fairly strict -- must be the exact length, uppercase, etc
// as the assumption is that the input reading function would have done data cleaning
function isValidUin(input) {
  if (input.length !== 9) {
    return false;
  }
  let parts = input.match(uinRegex);
  if (parts === null) {
    // failed to match the regex at all
    return false;
  }

  return parts[3] === uinChecksum(parts[1], parts[2]);
}

function render() {
  // step 1 table
  $('#render-import-totalRead').html(data.inputStats.totalRead);
  $('#render-import-totalInputUins').html(data.inputUin.size);
  $('#render-import-totalDuplicateUins').html(data.inputStats.duplicate);
  $('#render-import-totalInvalidUins').html(data.inputStats.invalid);

  // step 2 button
  if (data.inputUin.size === 0) {
    $('#trigger-process').prop('disabled', true);
  } else {
    $('#trigger-process').prop('disabled', false);
  }

  // step 2 table
  $('#render-retrieve-progress').html(`${data.inputUin.size - data.stats['retrieved']} of ${data.inputUin.size}`);
  $('#render-positive-test').html(data.stats['positiveTestResults']);
  $('#render-negative-test').html(data.stats['negativeTestResults']);
  $('#render-pending-test').html(data.stats['pendingTestResults']);
  $('#render-invalid-test').html(data.stats['invalidTestResults']);
  $('#render-no-test').html(data.stats['noTestResults']);
  $('#render-retrieve-countdown').html(data.progress);

  // step 3 button
  if (data.stats['retrieved'] === 0) {
    $('#export-file-button').prop('disabled', true);
  } else {
    $('#export-file-button').prop('disabled', false);
  }

  // error outputs
  let outputMessages = data.logs.join("\n");
  $('#render-errors').html(outputMessages);
}

function calculateStats() {
  let stats = {
    retrieved: 0,
    noTestResults: 0,
    positiveTestResults: 0,
    negativeTestResults: 0,
    pendingTestResults: 0,
    invalidTestResults: 0,
    withTestResults: 0
  };

  data.inputUin.forEach(function(uinData) {
    if (uinData['data']['results'].length === 0) {
      stats['noTestResults'] += 1
    } else {
      switch (uinData['data']['results'][0]['result']) {
        case "PENDING":
          stats['pendingTestResults'] += 1;
          break;
        case "INVALID":
          stats['invalidTestResults'] += 1;
          break;
        case "POSITIVE":
          stats['positiveTestResults'] += 1;
          break;
        case "NEGATIVE":
          stats['negativeTestResults'] += 1;
          break;
      }
    }
  });

  stats['withTestResults'] = stats['pendingTestResults'] + stats['invalidTestResults'] + stats['positiveTestResults'] + stats['negativeTestResults'];
  stats['retrieved'] = stats['withTestResults'] + stats['noTestResults'];
  data.stats = stats;
}

class ApiClient {
  constructor(base_url, key, concurrent_limit) {
    this.client = this.setupClient(base_url, key, concurrent_limit);

    // queue of objects, each containing the config (request) and the
    // corresponding promise resolve function
    this.queue = [];
    // accounting for currently in-flight requests
    this.current = [];
  }

  queueLength() {
    return this.queue.length;
  }

  setupClient(base_url, key, concurrent_limit) {
    const client = axios.create({
      baseURL: base_url,
      timeout: 20000,
      headers: { Authorization: `Bearer ${key}` },
      // ensure don't hit follow redirects issue by disabling it: https://github.com/axios/axios/issues/3217
      maxRedirects: 0,
    });
  
    const dispatch = () => {
      // dispatch only if possible
      if (this.current.length < concurrent_limit && this.queue.length > 0) {
        let item = this.queue.shift();
  
        // return the request back to axios for execution
        item.resolve(item.config);
  
        this.current.push(item);
      }
    }
  
    client.interceptors.request.use(config => {
      // create and return a new promise containing the config, the config is only
      // released back to axios for execution by the dispatch logic
      return new Promise(resolve => {
        this.queue.push({ config: config, resolve: resolve });
        dispatch();
      });
    }, error => {
      return Promise.reject(error);
    });
  
    client.interceptors.response.use(response => {
      this.current.shift();
      dispatch();
  
      return response;
    }, error => {
      // do exactly the same thing as non-errors, because need to do the same
      // accounting to return the concurrent quota and dispatch next
      this.current.shift();
      dispatch();
  
      return Promise.reject(error);
    });
  
    return client;
  }

  get(config) {
    return this.client.get(config);
  }
}

// 130 seconds for 80k at 100 concurrent limit, on local nginx 404s
// 92 sec for 1k at 20 concurrent, on staging
// 83 sec for 1k at 100 concurrent, on staging => probably concurrency limiting
function retrieveTestResults(base_url, key) {
  // update any UINs that require re-retrieval

  // if get an authentication error, stop everything early, do not continue
  let stopEarly = false;

  let client = new ApiClient(base_url, key, 20);
  let requests = [];

  const startTimestamp = $('#start_timestamp').val().trim();

  data.inputUin.forEach(function(uinData, key) {
    if (stopEarly) {
      return;
    }

    // skip UINs that already have a response
    if ('data' in uinData) {
      return;
    }

    let queryUrl = '';
    if (uinData['idType'] === 'uin') {
      queryUrl = `/results/patient?uin=${uinData['uin']}`;
    } else if (uinData['idType'] === 'passport') {
      queryUrl = `/results/patient?uin=${uinData['uin']}&uin_country_of_issue=${uinData['nationality']}`;
    }

    if (startTimestamp !== '') {
      queryUrl += `&start_timestamp=${startTimestamp}`
    }

    let p = client.get(queryUrl)
        .then(response => {
          // check status? or assume must be 2xx since not error?
          uinData['data'] = response['data'];

          if (client.queueLength() % 100 === 0) {
            data.progress = client.queueLength();
            render();
          }
        })
        .catch(error => {
          if (error.response) {
            // non-2xx response
            if (error.response.data.message === 'Authentication failed.') {
              data.logs.push('ERROR: Authentication failure encountered, stopping early');
              stopEarly = true;
            } else {
              data.logs.push('ERROR: Some unexpected response code');
              data.logs.push(error.response.headers);
              data.logs.push(error.response.data);
            }
          } else if (error.request) {
            // no response
            data.logs.push(`ERROR: No response ${error.request}`);
          } else {
            data.logs.push(`ERROR: Unexpected error from client ${error.message}`);
          }

          if (client.queueLength() % 100 === 0) {
            data.progress = client.queueLength();
            render();
          }
        });
    requests.push(p);

    if (client.queueLength() % 100 === 0) {
      data.progress = client.queueLength();
      render();
    }
  });

  // return a promise for all requests to complete
  return Promise.allSettled(requests);
}

function process() {
  let base_url = $('#api-baseurl').val();
  if (base_url.match(baseurlRegex) === null) {
    // base url is not valid, don't attempt to do anything
    data.logs.push('ERROR: API Base URL is invalid format');
    render();
    return;
  }

  let key = $('#api-key').val();
  if (key.match(keyRegex) === null) {
    // key is not valid, don't attempt to do anything
    data.logs.push('ERROR: API Key is invalid format');
    render();
    return;
  }

  // update any UINs that require re-retrieval
  retrieveTestResults(base_url, key)
  .then(() => {
    // re-calculate the counts based on all the collected results once
    calculateStats();
    render();
  });
}

function exportExcelFile(mode) {
  const workbook = new ExcelJS.Workbook();

  let statusSheet = workbook.addWorksheet('test_results');

  const remainingColumnHeaders = data.inputColumnHeaders.reduce((acc, header, i) => {
    // Ignore uin, nationality and passport
    if (i === 0 || i === 1 || i === 2 || i === 3) {
      return acc;
    }
    acc.push({ header: header, key: header, width: 15 });
    return acc;
  }, [])

  statusSheet.columns = [
    { header: 'UIN', key: 'uin', width: 15 },
    { header: 'Nationality', key: 'nationality', width: 10 },
    { header: 'Passport Number', key: 'passport', width: 15 },
    { header: 'Covid test result', key: 'result', width: 15 },
    { header: 'Swab Reason', key: 'swab_reason', width: 15 },
    { header: 'Produced At', key: 'produced_at', width: 15 },
    ...remainingColumnHeaders
  ]

  // go ahead to set the obj[key] for everything, as they will be ignored depending on the
  // statusSheet.columns config above
  data.inputUin.forEach(function(uinData, uin) {
    let obj = {
      uin: uinData['uin'],
      nationality: uinData['nationality'],
      passport: uinData['passport'],
      ...uinData['otherInfo']
    }

    const results = uinData['data']['results'];
    
    if (results.length === 0) {
      obj['result'] = "NO RESULT";
    } else {
      obj['result'] = results[0]['result'];
      obj['swab_reason'] = results[0]['swab_reason']
      obj['produced_at'] = new Date(results[0]['produced_at']).toLocaleString()
    }

    statusSheet.addRow(obj);
  });

  let now = new Date();
  workbook.xlsx.writeBuffer()
    .then(buffer => {
      let blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      let objUrl = window.URL.createObjectURL(blob);
      let anchor = document.createElement('a');
      anchor.href = objUrl;
      anchor.download = `status-${now.getFullYear()}${('0' + (now.getMonth() + 1)).slice(-2)}${('0' + now.getDate()).slice(-2)}-${('0' + now.getHours()).slice(-2)}${('0' + now.getMinutes()).slice(-2)}${('0' + now.getSeconds()).slice(-2)}`;
      anchor.click();
      window.URL.revokeObjectURL(objUrl);
    });
}

// given a File object, parse it, validate it, and insert any new valid IDs into the inputUin state
async function importExcelFile(file) {
  let buffer = await file.arrayBuffer();
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer);

  workbook.eachSheet(function(worksheet, sheetId) {
    // some users don't know about their hidden worksheets, ignore those worksheets
    if (worksheet.state !== 'visible') {
      data.logs.push(`WARNING: file '${file.name}' worksheet '${worksheet.name}' sheet is hidden, ignoring`);
      return;
    }

    // confirm that at least A1 is as expected
    let firstCell = worksheet.getCell('A1');
    if (firstCell.value !== 'UIN') {
      data.logs.push(`WARNING: file '${file.name}' worksheet '${worksheet.name}' cell A1 does not contain 'UIN'`);
      return;
    }

    data.inputColumnHeaders = worksheet.getRow(1).values

    worksheet.eachRow(function (row, rowNumber) {
      // skip first row, checked above to be header row
      if (rowNumber === 1) {
        return;
      }

      // skip row if the A and B cells are empty
      if (row.getCell('A').text === '' && row.getCell('B').text === '') {
        return;
      }

      // only count rows where the A and B cells are non-empty
      data.inputStats.totalRead++;

      // cleaning the ID value to only retain the expected chars
      // use .text instead of .value, to get the string value of the cell.
      // otherwise, if there's rich text formatting within the cell, it becomes a richtext
      // object instead of a string.

      let idString = row.getCell('A').text.toUpperCase();
      idString = idString.replace(/[^A-Z0-9]/g, '');

      const nationality = row.getCell('B').text.toUpperCase();

      let passportString = row.getCell('C').text.toUpperCase();
      passportString = passportString.replace(/[^A-Z0-9]/g, '');

      let idType = 'uin';
      if (idString.length === 0) {
        idType = 'passport';
      }

      // nothing left after cleaning is considered read but invalid, not empty/unread
      if (idType === 'uin' && !isValidUin(idString)) {
        data.inputStats.invalid++;
        data.logs.push(`Invalid UIN: worksheet '${worksheet.name}' cell '${row.getCell('A').address}' value '${row.getCell('A').text}'`);
        return;
      } else if (idType === 'passport' && nationality.length !== 2) {
        data.inputStats.invalid++;
        data.logs.push(`Invalid Passport Number or Country Code: worksheet '${worksheet.name}' cell '${row.getCell('B').address}' value '${row.getCell('B').text}' '${row.getCell('C').text}'`);
        return;
      }

      let hashKey = idType === 'passport' ? `${idType}-${nationality}-${passportString}` : `${idType}-${idString}-${nationality}-${passportString}`;
      
      // insert into the map, which deals with de-duplication
      if (!data.inputUin.has(hashKey)) {
        const remainingInformation = data.inputColumnHeaders.reduce((acc, header, i) => {
          // Ignore uin, nationality and passport
          if (i === 0 || i === 1 || i === 2 || i === 3) {
            return acc;
          }
          acc[header] = row.getCell(i).text;
          return acc;
        }, {})

        data.inputUin.set(hashKey, { idType: idType, uin: idString, nationality: nationality, passport: passportString, otherInfo:  { ...remainingInformation } });
      } else {
        data.inputStats.duplicate++;
        data.logs.push(`Duplicate ID: worksheet '${worksheet.name}' cell '${row.getCell('A').address}' value '${row.getCell('A').text}' '${row.getCell('B').text}' '${row.getCell('C').text}`);
      }
    });
  });
}

function importExcelFiles() {
  // FileList
  let files = $('#import-files').prop('files');
  if (files === undefined) {
    // no file(s) selected, ignore
    return;
  }

  for (let i = 0; i < files.length; i++) {
    let file = files[i];
    importExcelFile(file).then(function() {
      render();
    });
  }
}

$(document).ready(function() {
  $('#import-file-button').click(function(event) {
    importExcelFiles();
  });

  $('#trigger-process').click(function(event) {
    process();
  });

  $('#export-file-button').click(function(event) {
    exportExcelFile($('#export-mode').val());
  });
});
