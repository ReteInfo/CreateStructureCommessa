using System;
using System.IO;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;


namespace Company.Function
{
    public static class CreateTeamStructure
    {
        [FunctionName("CreateTeamStructure")]
        public static async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Anonymous, "get", "post", Route = null)] HttpRequest req,ILogger log)
        {
            /*log.LogInformation("C# HTTP trigger function processed a request.");

            string name = req.Query["name"];

            string requestBody = await new StreamReader(req.Body).ReadToEndAsync();
            dynamic data = JsonConvert.DeserializeObject(requestBody);
            name = name ?? data?.name;

            string responseMessage = string.IsNullOrEmpty(name)
                ? "This HTTP triggered function executed successfully. Pass a name in the query string or in the request body for a personalized response."
                : $"Hello, {name}. This HTTP triggered function executed successfully.";

            return new OkObjectResult(responseMessage);*/
            try
            {
              
              string requestBody = await new StreamReader(req.Body).ReadToEndAsync();
              var input = JsonConvert.DeserializeObject<Team>(requestBody);
              //var task = new Task() { TaskDescription = input.TaskDescription };
              //Items.Add(task);
              //GraphServiceClient graphClient = Authentication.auth();
              //return new OkObjectResult(input);
              return null;
            }
            catch (System.Exception e)
            {
                return new ObjectResult(e.Message);
                /*var result = new ObjectResult(e.Message);
                result.StatusCode = StatusCodes.Status401Unauthorized;
                return result;*/
            }
        }
    }
}
