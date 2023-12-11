#!/bin/sh

cd /Users/jeffneal/projects/dotnet_source/MyWebApi/MyWebApi

podman stop $(podman ps -q)
dotnet clean

dotnet publish "MyWebApi.csproj" -f net7.0 -c Release

podman build -t mywebapi .   

# Bind to localhost
podman run -p 8080:80 -d mywebapi

# Bind to all interfaces
#podman run -p 0.0.0.0:8080:80 -d mywebapi

podman ps

#podman logs <container id>

#podman stop <container id>
#podman stop $(podman ps -q)

# Browser - http://localhost:8080/swagger/index.html

