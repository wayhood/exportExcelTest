#!/bin/make
all:
	export GO111MODULE=on
	export GOPROXY=https://goproxy.io,direct
	export GOSUMDB=off
	export GOOS=linux
	export GOARCH=amd64
	export CGO_ENABLED=0
	@rm -rf ../bin/golang-excel > /dev/null
	go build -ldflags "-s -w" -o ../bin/golang-excel main.go
