dotnet new classlib --name ServerCommandProxy
cd ServerCommandProxy
Remove-Item Class1.cs
Copy-Item 'C:\Program Files\PaperCut MF\server\examples\webservices\csharp\ServerCommandProxy.cs'
dotnet add package Kveer.XmlRPC --version 1.1.1
dotnet build --configuration Release
cd ..