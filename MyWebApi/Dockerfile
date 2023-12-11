FROM mcr.microsoft.com/dotnet/aspnet:7.0 AS base
WORKDIR /app
EXPOSE 80

#FROM mcr.microsoft.com/dotnet/sdk:7.0 AS build
#WORKDIR /src
#COPY ["/MyWebApi.csproj", "./MyWebApi/"]
#RUN dotnet restore "./MyWebApi/MyWebApi.csproj"
#COPY . .
#WORKDIR "/src/MyWebApi"
#RUN dotnet build "MyWebApi.csproj" -c Release -o /app/build

ADD bin/Release/net7.0/publish/ /app/publish/
ADD ./Keys/ /app/publish/Keys/
ENV GOOGLE_APPLICATION_CREDENTIALS=/app/publish/Keys/cybersecurity-class-404013-6f8510532d40.json
#FROM build AS publish
#RUN dotnet publish "MyWebApi.csproj" -c Release -o /app/publish

FROM base AS final
WORKDIR /app/publish
#COPY --from=publish /app/publish .
#COPY /app/publish/ .
#ENTRYPOINT ["dotnet", "MyWebApi.dll"]
CMD [ "dotnet", "MyWebApi.dll" ]