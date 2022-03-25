# covid-results-checker

Check the covid test results of a batch of individuals as identified by their UIN.

## Development

```

# run a http webserver in the root
docker run --rm -it -p 8080:80 -v $(pwd):/usr/share/nginx/html nginx:alpine
```
