using System;
using System.Drawing;
using System.Management.Automation;
using System.Net.Http.Headers;
using System.Security.Policy;
using System.Text;

namespace SendTeamsMessageUsingIncomingWebhook
{

    [Cmdlet(VerbsCommon.Push, "TeamsMesageUsingWebHook")]
    [OutputType(typeof(ResultData))]
    public class SendTeamsMeesageUsingWebHook:PSCmdlet
    {
        [Parameter(Mandatory = true, Position = 0, HelpMessage = "https://learn.microsoft.com/en-us/microsoftteams/platform/webhooks-and-connectors/how-to/add-incoming-webhook?tabs=dotnet", ValueFromPipeline = false)]
        [ValidateNotNullOrEmpty()]
        public  string IncomingWebHookUrl { get; set; }

        [Parameter(Mandatory = true, Position = 1, HelpMessage = "Message that you want to send to your Teams Channel", ValueFromPipeline = false)]
        [ValidateNotNullOrEmpty()]
        public string MessageToSend { get; set; }

        [Parameter(Mandatory = false, Position = 2, HelpMessage = "Provide complete path to JPEG file, pleasse dont upload image size more than 28 KB as total payload size is  28 KB", ValueFromPipeline = false)]
        [ValidateNotNullOrEmpty()]

        public string JpgImageFilePath { get; set; } = "default";


        public bool jpgImagePath { get; set; }=false;

        public StringBuilder adaptiveCard = new();
        protected override void   ProcessRecord()
        {
            
                var adaptiveJson = @"{

                              ""type"": ""message"",

                              ""attachments"": [

                                {

                                  ""contentType"": ""application/vnd.microsoft.card.adaptive"",

                                  ""content"": {

                                    ""type"": ""AdaptiveCard"",

                                    ""body"": [

                                      {

                                        ""type"": ""TextBlock"",

                                        ""text"": ""1234"",


                                      },

{

                                        ""type"": ""Image"",
                                        ""url"":""data""
                                      


                                      }


                                    ],

                                    ""$schema"": ""http://adaptivecards.io/schemas/adaptive-card.json"",

                                        ""version"": ""1.0""

                                  }

                                }

                              ]

                            }";

          
          
            adaptiveCard.Append(adaptiveJson);

            if(JpgImageFilePath!="default")
            {
                if (File.Exists(JpgImageFilePath))
                {

                    jpgImagePath = true;
                    var base64string = ConvertImageToBase64(JpgImageFilePath);
                    if (!string.IsNullOrEmpty(base64string))
                    {
                        adaptiveCard.Replace("data", base64string);
                        adaptiveCard.Replace("1234", MessageToSend);


                        WriteObject(SendTeamsMessage(adaptiveCard.ToString(), IncomingWebHookUrl, jpgImagePath));
                    }
                    else
                    {
                        WriteObject(new ResultData { ReturnCode = 500, ReturnMessage = $"Image to Base64  conversion probably failed!!" });



                    }

                }
                else
                {
                    WriteObject(new ResultData { ReturnCode = 300, ReturnMessage = $"Absoulte path JPEF file is required if you need to upload an image of size less than 28KB" });

                }

            }
            else
            {

                adaptiveCard.Replace("data", "data:image/jpeg;base64, iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=");
                adaptiveCard.Replace("1234", MessageToSend);
                WriteObject(SendTeamsMessage(adaptiveCard.ToString(), IncomingWebHookUrl, jpgImagePath));
            }




        }


        public static ResultData SendTeamsMessage(string message, string webhook, bool messageType)
        {

           //Console.WriteLine(message);

            if (!string.IsNullOrEmpty(message))
            {

          
                try
                {
                    var client = new HttpClient();
                    client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

                    var content = new StringContent(message, System.Text.Encoding.UTF8, "application/json");
                    var response = client.PostAsync(webhook, content).ConfigureAwait(true).GetAwaiter();
                    if (response.GetResult().StatusCode == System.Net.HttpStatusCode.OK)
                    {
                        var msg = messageType ? "Message sent with Image!!" : "Message sent without Image!!";
                        return new ResultData { ReturnCode = 0, ReturnMessage = $"Success!! {msg}" };

                    }
                    else
                    {
                        return new ResultData { ReturnCode = 100, ReturnMessage = $"Oh! Something Went Wrong,  {response.GetResult().StatusCode} " };

                    }

                }
                catch (Exception ex)
                {

                    return new ResultData { ReturnCode = -2, ReturnMessage = ex.Message };
                }


            }
            else
            {
                return new ResultData { ReturnCode = -200, ReturnMessage = "Please check your inputs,Something went Wrong" };

            }



        }


        public static string ConvertImageToBase64(string imagepath)
        {
            try
            {
                var jpgimage = File.ReadAllBytes(imagepath);
                var base64output = Convert.ToBase64String(jpgimage,Base64FormattingOptions.None);

               
                return String.Format($"data:image/png;base64,{base64output}", base64output);
            }
            catch(Exception ex)
            {
                return string.Empty;
            }
        }

    }



}