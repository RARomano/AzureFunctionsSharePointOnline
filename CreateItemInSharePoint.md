Esse exemplo mostra como utilizar o Azure Functions para criar um item no SharePoint baseado em uma requisição REST (Javascript)

## Preparando a solução
1- Clone o repositório

2- Crie uma lista no SharePoint

3- Crie um add-in no SharePoint (https://tenant.sharepoint.com/_layouts/appregnew.aspx)

4- Conceda as permissões necessárias nessa lista (https://tenant.sharepoint.com/_layouts/appinv.aspx)
```xml
<AppPermissionRequests AllowAppOnlyPolicy="true">
    <AppPermissionRequest Scope="http://sharepoint/content/sitecollection/web/list" Right="Manage" />
</AppPermissionRequests>
```

5- Altere o arquivo [config.js](https://github.com/RARomano/AzureFunctionsSharePointOnline/blob/master/SPFunctionJS/config.js) com os dados gerados nas etapas anteriores:
```javascript
module.exports = {
  // client ID gerado
  clientId: '',
  // client Secret gerado
  clientSecret: '',
  sharePointPrinciple:'00000003-0000-0ff1-ce00-000000000000',
  wwwauthenticate:'www-authenticate',
  clientSvcUrl:'/_vti_bin/client.svc',
  authenticationUrl:'https://accounts.accesscontrol.windows.net/%s/tokens/OAuth/2',

  // site url (https://tenant.sharepoint.com)
  siteUrl: '',
  // lista criada (Ex: Teste)
  listUrl: ''
};
```


## Configurando Continuous Integration
1- Entre no **Function App Settings** e depois clique em **Configure continuous integration**

![image](https://cloud.githubusercontent.com/assets/12012898/22079529/5da1d9b8-dda3-11e6-8bd1-b2220b57ed88.png)

2- Clique em **Setup**

![image](https://cloud.githubusercontent.com/assets/12012898/22079586/925a8ba0-dda3-11e6-9dad-d3c78df5337e.png)

3- Escolha os dados do repositório
O nome de cada pasta que você criar no seu repositório será utilizado como nome da função.

![image](https://cloud.githubusercontent.com/assets/12012898/22079622/ac9d2428-dda3-11e6-878d-991a362a65a0.png)

## Executando a função
1- Pegue a url da função e invoke-a
```javascript
$(function() {
    $.get('https://rrtestfunctions.azurewebsites.net/api/SPFunctionJS?name='+_spPageContextInfo.userLoginName, function(result) { 
        console.log(result); 
    });
});
```

2- Um item deverá ser criado no site que você configurou.

![image](https://cloud.githubusercontent.com/assets/12012898/22079837/5db21390-dda4-11e6-9d57-b243e41d294c.png)

![image](https://cloud.githubusercontent.com/assets/12012898/22079856/6edba226-dda4-11e6-8a07-fffb0d585e2f.png)
