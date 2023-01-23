## Dependencies

1. [OpenPop.NET](https://hpop.sourceforge.net/)
2. [User Secrets](https://learn.microsoft.com/en-us/aspnet/core/security/app-secrets?view=aspnetcore-7.0&tabs=windows#secret-manager)

## Config for User Secrets

There's two ways of creating the file and writing inside of it:

<details>
  <summary>Using Visual Studio</summary>
  <p>    
    
    1. Right-click on the Solution Explorer
    2. Click on the option Manage User Secrets    
    
  </p>
</details>

<details>
  <summary>Using .NET CLI</summary>
  
  First you use the following command to create the file where it will be stored
  
  <p>    
    
   ```PowerShell
   dotnet user-secrets init
   ```
    
  </p>
  
  And adding another command to create a secret value inside the recent created file:
  
  <p>    
    
   ```PowerShell
   dotnet user-secrets set "MySecret" "12345"
   ```
    
  </p>
  
  After that, the file will look someting like this:
  
  <p>    
    
   ```json
  {
      "MySecret": "12345"
  }
  ```
    
  </p>
</details>

In the end your _**secrets.json**_ file need to follow the structure below, where you need to edit all the values accordingly to you necessity

```json
{
  "AuthenticationData": {
    "Email": "string",
    "Password": "string",
    "Hostname": "string",
    "Port": "int",
    "UseSSL": "bool"
  },
  "EmailReceived": {
    "Address": "string",
    "Attachment1": "string",
    "Attachment2": "string"
  },
  "Others": {
    "TargetPath": "string"
  }
}
```
