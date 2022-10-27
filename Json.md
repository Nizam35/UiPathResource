
## Get Specific item vlaue from Json using Select Token

``` vb.net

'Example
{
  "@odata.context": "https://devorch.bupa.com.sa/odata/$metadata#QueueDefinitions",
  "@odata.count": 4,
  "value": [
    {
      "Name": "Complaints",
      "Description": "",
      "MaxNumberOfRetries": 0,
      "AcceptAutomaticallyRetry": false,
      "EnforceUniqueReference": false,
      "SpecificDataJsonSchema": null,
      "OutputDataJsonSchema": null,
      "AnalyticsDataJsonSchema": null,
      "CreationTime": "2021-04-07T13:20:18.517Z",
      "ProcessScheduleId": null,
      "SlaInMinutes": 0,
      "RiskSlaInMinutes": 0,
      "ReleaseId": null,
      "Id": 10015
    },
    {
      "Name": "SecondPreauthIDs",
      "Description": "",
      "MaxNumberOfRetries": 1,
      "AcceptAutomaticallyRetry": true,
      "EnforceUniqueReference": false,
      "SpecificDataJsonSchema": null,
      "OutputDataJsonSchema": null,
      "AnalyticsDataJsonSchema": null,
      "CreationTime": "2021-06-07T11:48:21.943Z",
      "ProcessScheduleId": null,
      "SlaInMinutes": 0,
      "RiskSlaInMinutes": 0,
      "ReleaseId": null,
      "Id": 10037
    },
    {
      "Name": "BankingDetailsUpdate_Queue",
      "Description": "To update the Banking Details",
      "MaxNumberOfRetries": 1,
      "AcceptAutomaticallyRetry": true,
      "EnforceUniqueReference": false,
      "SpecificDataJsonSchema": null,
      "OutputDataJsonSchema": null,
      "AnalyticsDataJsonSchema": null,
      "CreationTime": "2021-11-02T07:55:47.94Z",
      "ProcessScheduleId": null,
      "SlaInMinutes": 0,
      "RiskSlaInMinutes": 0,
      "ReleaseId": null,
      "Id": 10043
    },
    {
      "Name": "DataValidation_Queue",
      "Description": "testing",
      "MaxNumberOfRetries": 1,
      "AcceptAutomaticallyRetry": true,
      "EnforceUniqueReference": false,
      "SpecificDataJsonSchema": null,
      "OutputDataJsonSchema": null,
      "AnalyticsDataJsonSchema": null,
      "CreationTime": "2022-02-10T11:51:52.98Z",
      "ProcessScheduleId": null,
      "SlaInMinutes": 0,
      "RiskSlaInMinutes": 0,
      "ReleaseId": null,
      "Id": 10044
    }
  ]
}
in_Config("OrchestratorQueueName") = "DataValidation_Queue"

CInt(JObject.SelectToken("$.value[?(@.Name == '"+in_Config("OrchestratorQueueName").ToString+"')]").Item("MaxNumberOfRetries"))

```
