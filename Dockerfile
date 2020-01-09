FROM mcr.microsoft.com/dotnet/core/aspnet:3.1.0 AS base
WORKDIR /app
EXPOSE 80

FROM mcr.microsoft.com/dotnet/core/sdk:3.1.100 AS build
WORKDIR /src
COPY SharePointAPI.sln ./
COPY SharePointAPI.csproj ./
RUN dotnet restore "SharePointAPI.csproj" 
COPY . .
RUN dotnet build "SharePointAPI.csproj" -c Release -o /app

FROM build AS publish
RUN dotnet publish "SharePointAPI.csproj" -c Release -o /app

FROM base AS final
WORKDIR /app
COPY --from=publish /app .
ENTRYPOINT ["dotnet", "SharePointAPI.dll"]