# AzureFunctionsSharePointOnline
Repositório para demonstrar como utilizar o Azure Functions para fazer processamento do lado do Servidor em um arquitetura Serverless e invocando essas funções diretamente do SharePoint.

Esse recurso do Azure, o Azure Functions, é bem interessante pois ele permite que o seu código seja executado sem precisar de um servidor inteiro rodando pra ele (Serverless Architecture). Veja mais em: https://azure.microsoft.com/pt-br/services/functions/

Isso é bem útil em cenários que precisamos realizar alguns processamentos mais sensíveis e que o código não pode estar disponível para o usuário.

Em uma SharePoint-Hosted add-ins, o código roda no contexto do browser do usuário, ou seja, se for um usuário avançado ele pode debugar ou até alterar o código de forma mal-intencionada. Mas, utilizado em conjunto com o Azure Functions, você pode contornar essa "limitação".

## Pré-requisitos
- O site de onde a função será invocada deverá ter a biblbioteca jQuery
- Azure functions app

## Configurando o Azure functions app
1- Vá no portal azure https://portal.azure.com e clique em **New**, depois **Compute** e depois em **Function app**
![image](https://cloud.githubusercontent.com/assets/12012898/22021524/0536ce0e-dca5-11e6-8e86-257c4775646b.png)

2- Preencha os dados e clique em **Create**
![image](https://cloud.githubusercontent.com/assets/12012898/22021585/46f448c6-dca5-11e6-927f-a0dda5947aff.png)

3- Abra o App recém criado e clique em **New function**
![image](https://cloud.githubusercontent.com/assets/12012898/22021630/7b0a592a-dca5-11e6-8c3e-6bf9161ac464.png)

4- Escolha **HTTP Trigger - Javascript**, dê um nome para a função e mude o Authorization Level para **Anonymous** (nesse exemplo não será necessário autorização para a nossa API, mas certamente nas necessidades reais do dia-a-dia, isso deve ser mudado).
![image](https://cloud.githubusercontent.com/assets/12012898/22021682/b34600aa-dca5-11e6-8cd8-a9a2777a6389.png)

5- Após criar, vá em **Function app settings** e depois clique em **Configure CORS**
![image](https://cloud.githubusercontent.com/assets/12012898/22021754/10e7ef8e-dca6-11e6-8819-44dc8258fcd8.png)

6- Adicione o site de onde essa função será chamada e clique em **OK**
![image](https://cloud.githubusercontent.com/assets/12012898/22021778/2e8a2e6c-dca6-11e6-9ab8-f8533ce82df5.png)

7- Clique na função e altere o código dela conforme a sua necessidade
> Nesse exemplo, utilizamos Javascript (Node) mas pode ser utilizado C#. Além disso, estou fazendo um cenário bem simples para demonstração das funcionalidades. Em um próximo artigo, vou demonstrar como gravar dados no SharePoint.

![image](https://cloud.githubusercontent.com/assets/12012898/22021808/56ef3c9e-dca6-11e6-8826-afbaf779abfc.png)

## Chamando essa função no site SharePoint
Para efeitos de demonstração, eu só estou invocando o código passando o usuário logado e imprimindo o retorno no console.
```javascript
$(function() {
	$.get('https://rrtestfunctions.azurewebsites.net/api/SharePointTest?name='+_spPageContextInfo.userLoginName, function(result) { 
		console.log(result); 
	});
});
```

Para testes, você pode rodar essa código diretamente no console do chrome:
![image](https://cloud.githubusercontent.com/assets/12012898/22022184/06c8069a-dca8-11e6-978b-7214b71041e5.png)
