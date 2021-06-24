# Quick POC for PPT Gen using express, docker and redis

Requirements: [Docker Community Edition](https://www.docker.com/community-edition)

To start the app run: `docker compose up`.

It will then be started on port 3000.

# Endpoints

## Hello World

```sh
curl http://localhost:3000
```

## Generate powerpoint
```sh
curl http://localhost:3000/powerpoint
```
