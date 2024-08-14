FROM mcr.microsoft.com/dotnet/sdk:8.0 AS build
WORKDIR dockerAPI

EXPOSE 80
EXPOSE 5024

COPY ./*.csproj ./
RUN dotnet restore

COPY . .
RUN dotnet publish -c Release -o out

# Build runtime image
FROM mcr.microsoft.com/dotnet/aspnet:8.0
WORKDIR /dockerAPI
COPY --from=build /dockerAPI/out .
ENTRYPOINT ["dotnet", "APIDocker.dll"]