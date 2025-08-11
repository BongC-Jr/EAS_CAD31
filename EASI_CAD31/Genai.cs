using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using Newtonsoft.Json;
using Newtonsoft.Json.Serialization;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Threading;



namespace EASI_CAD31
{
    public class GeminiClient
    {
        private readonly string _apiKey;
        private readonly string _model;
        private readonly HttpClient _httpClient;

        //Models: gemini-2.0-flash gemini-2.5-pro
        public GeminiClient(string apiKey, string model = "gemini-2.0-flash")
        {
            _apiKey = apiKey;
            _model = model;
            _httpClient = new HttpClient();
        }

        public async Task<string> GenerateContentAsync(string prompt)
        {
            string currentPath = Application.DocumentManager.MdiActiveDocument.Database.Filename;
            string currentDirectory = "";
            if (!string.IsNullOrEmpty(currentPath))
            {
                currentDirectory = System.IO.Path.GetDirectoryName(currentPath);
            }
            DataGlobal.convofilepath = currentDirectory;


            var contentsList = new List<dynamic>
            {
                new { role = "user", parts = new[] { new { text = $"The current directory with filename of the drawing " +
                                                                  $"is {currentPath.ToString()}. If you are asked for the "+
                                                                  $"directory or path or folder, just give the directory without "+
                                                                  $"the filename. If you are asked for the filename, just give "+
                                                                  $"the filename. Ok?" } } },
                new { role = "assistant", parts = new[] { new { text = "I confirm." } } },
            };

            IList<IList<object>> iioConvoData = DataGlobal.trainingConversation;
            foreach (var row in iioConvoData)
            {
                string role = row[0].ToString();
                string text = row[1].ToString(); // Assuming the second column is the text content
                 
                contentsList.Add(new
                {
                    role = role,
                    parts = new[] { new { text = text } }
                });
            }


            string convofilePath = Path.Combine(DataGlobal.convofilepath, "conversation.txt");
            if (File.Exists(convofilePath))
            {
                string[] convoLines = File.ReadAllLines(convofilePath);

                dynamic currentUserTurn = null; // Stores the 'User' part of a potential pair

                foreach (string line in convoLines)
                {
                    string trimmedLine = line.Trim(); // Trim whitespace from the line itself
                    if (string.IsNullOrWhiteSpace(trimmedLine))
                    {
                        // If we encounter a blank line, it just means the previous pair is finished.
                        // If currentUserTurn is not null here, it means a User turn was left unpaired.
                        // We just let it get discarded when the loop continues.
                        currentUserTurn = null; // Reset for a new potential pair
                        continue;
                    }

                    string role = "";
                    string text = "";

                    if (trimmedLine.StartsWith("user: ", StringComparison.OrdinalIgnoreCase))
                    {
                        role = "user";
                        text = trimmedLine.Substring("user: ".Length).Trim();

                        // If there was an unpaired User turn previously, discard it
                        if (currentUserTurn != null)
                        {
                            // Optionally log: ed.WriteMessage($"\nWarning: Discarding unpaired User turn from history: '{currentUserTurn.parts[0].text}'");
                        }
                        // Start a new User turn
                        currentUserTurn = new { role = role, parts = new[] { new { text = text } } };
                    }
                    else if (trimmedLine.StartsWith("assistant: ", StringComparison.OrdinalIgnoreCase))
                    {
                        role = "assistant";
                        text = trimmedLine.Substring("assistant: ".Length).Trim();

                        // This is an Assistant turn. It must be paired with a preceding User turn.
                        if (currentUserTurn != null)
                        {
                            // Add the stored User turn
                            contentsList.Add(currentUserTurn);
                            // Add the current Assistant turn
                            contentsList.Add(new { role = role, parts = new[] { new { text = text } } });

                            // Reset for the next pair
                            currentUserTurn = null;
                        }
                        else
                        {
                            // This Assistant turn is not preceded by a User turn. Exclude it.
                            // Optionally log: ed.WriteMessage($"\nWarning: Skipping unpaired Assistant turn: '{trimmedLine}'");
                        }
                    }
                    else
                    {
                        // Line does not start with "User:" or "Assistant:". Exclude it.
                        // If there was a pending User turn before this malformed line, it's also implicitly discarded.
                        currentUserTurn = null; // Discard any pending User turn as the sequence is broken
                                                // Optionally log: ed.WriteMessage($"\nWarning: Skipping malformed line in convo.txt: '{trimmedLine}'");
                    }
                }

                // After the loop, if currentUserTurn is not null, it means the last User turn
                // in the file did not have a corresponding Assistant turn. It will be excluded automatically
                // because it was never added to contentsList.
                // Optionally log: if (currentUserTurn != null) { ed.WriteMessage($"\nWarning: Last User turn in convo.txt was unpaired and excluded."); }

                // ed.WriteMessage($"\nLoaded conversation history from convo.txt."); // You can re-add this if you like.
            }
            else
            {
                // ed.WriteMessage($"\nWarning: conversation.txt not found at '{convofilePath}'. No conversation history loaded from file.");
            }

            //string convofilePath = Path.Combine(DataGlobal.convofilepath, "conversation.txt");
            //if (File.Exists(convofilePath))
            //{
            //    string[] convoLines = File.ReadAllLines(convofilePath);
            //    foreach (string line in convoLines)
            //    {
            //        if (string.IsNullOrWhiteSpace(line)) continue; // Skip empty lines
            //
            //        string role = "";
            //        string text = "";
            //
            //        if (line.StartsWith("User: ", StringComparison.OrdinalIgnoreCase))
            //        {
            //            role = "user";
            //            text = line.Substring("User: ".Length).Trim();
            //        }
            //        else if (line.StartsWith("Assistant: ", StringComparison.OrdinalIgnoreCase))
            //        {
            //            role = "assistant";
            //            text = line.Substring("Assistant: ".Length).Trim();
            //        }
            //        else
            //        {
            //            // Optional: Log a warning for lines that don't match the expected format
            //            //ed.WriteMessage($"\nWarning: Skipping malformed line in convo.txt: '{line}'");
            //            continue;
            //        }
            //
            //        contentsList.Add(new { role = role, parts = new[] { new { text = text } } });
            //    }
            //}

            contentsList.Add(new { role = "user", parts = new[] { new { text = prompt } } });

            var request = new
            {
                contents = contentsList.ToArray(),
                generationConfig = new { responseMimeType = "text/plain" }
            };


            //var request = new
            //{
            //    contents = new[]
            //    {
            //        new { role = "user", 
            //              parts = new[] { new { text = prompt } } }
            //    },
            //    generationConfig = new
            //    {
            //        responseMimeType = "text/plain"
            //    }
            //};

            var json = JsonConvert.SerializeObject(request);
            var content = new StringContent(json, Encoding.UTF8, "application/json");

            var url = $"https://generativelanguage.googleapis.com/v1beta/models/{_model}:generateContent?key={_apiKey}";
            var response = await _httpClient.PostAsync(url, content);
            response.EnsureSuccessStatusCode();

            var responseBody = await response.Content.ReadAsStringAsync();
            var responseData = JsonConvert.DeserializeObject<dynamic>(responseBody);

            return responseData.candidates[0].content.parts[0].text;
        }
    }

    public class AICommands
    {
       static Document actDoc = Application.DocumentManager.MdiActiveDocument;
       static Database aCurDB = actDoc.Database;

      [CommandMethod("CAI_")]
      public static async void CAI2()
      {
         //GetUrlResponseAsync().GetAwaiter().GetResult();
         //GetUrlResponse();
          PromptStringOptions psoUserMsg = new PromptStringOptions("\nQuestion: ");
          psoUserMsg.AllowSpaces = true;
          PromptResult prUserMsg = actDoc.Editor.GetString(psoUserMsg);
          if (prUserMsg.Status != PromptStatus.OK) return;
          string userMessage = prUserMsg.StringResult;
 
          actDoc.Editor.WriteMessage("\nPlease wait... Thinking...");
 
          Thread.Sleep(300); 
          // URL to send the POST request to
          string url = "http://127.0.0.1:7788/aibot/cai_d4h?v1=dfg84j567";
          
          string currentPath = Application.DocumentManager.MdiActiveDocument.Database.Filename;
          string currentDirectory = "";
          if (!string.IsNullOrEmpty(currentPath))
          {
              currentDirectory = System.IO.Path.GetDirectoryName(currentPath);
          }
          DataGlobal.convofilepath = currentDirectory;
          
          var contentsList = new List<dynamic>
          {
              new { role = "user", parts = new[] { new { text = $"The current directory with filename of the drawing " +
                                                                $"is {currentPath.ToString()}. If you are asked for the "+
                                                                $"directory or path or folder, just give the directory without "+
                                                                $"the filename. If you are asked for the filename, just give "+
                                                                $"the filename. Ok?" } } },
              new { role = "model", parts = new[] { new { text = "I confirm." } } },
          }; 
          

         IList<IList<object>> iioConvoData = DataGlobal.trainingConvCai2;
          
         foreach (var row in iioConvoData)
         {
            string role = row[0].ToString();
            string text = row[1].ToString(); // Assuming the second column is the text content

            contentsList.Add(new
            {
               role = role,
               parts = new[] { new { text = text } }
            });
         }


         string convofilePath = Path.Combine(DataGlobal.convofilepath, "conversation.txt");
         if (File.Exists(convofilePath))
         {
             string[] convoLines = File.ReadAllLines(convofilePath);
             
             dynamic currentUserTurn = null; // Stores the 'User' part of a potential pair
             
             foreach (string line in convoLines)
             {
                 string trimmedLine = line.Trim(); // Trim whitespace from the line itself
                 if (string.IsNullOrWhiteSpace(trimmedLine))
                 {
                     // If we encounter a blank line, it just means the previous pair is finished.
                     // If currentUserTurn is not null here, it means a User turn was left unpaired.
                     // We just let it get discarded when the loop continues.
                     currentUserTurn = null; // Reset for a new potential pair
                     continue;
                 }
                 
                 string role = "";
                 string text = "";
                 
                 if (trimmedLine.StartsWith("user: ", StringComparison.OrdinalIgnoreCase))
                 {
                     role = "user";
                     text = trimmedLine.Substring("user: ".Length).Trim();
                           
                     // If there was an unpaired User turn previously, discard it
                     if (currentUserTurn != null)
                     {
                         // Optionally log: ed.WriteMessage($"\nWarning: Discarding unpaired User turn from history: '{currentUserTurn.parts[0].text}'");
                     }
                     // Start a new User turn
                     currentUserTurn = new { role = role, parts = new[] { new { text = text } } };
                 }
                 else if (trimmedLine.StartsWith("model: ", StringComparison.OrdinalIgnoreCase))
                 {
                     role = "model";
                     text = trimmedLine.Substring("model: ".Length).Trim();
                     
                     // This is an Assistant turn. It must be paired with a preceding User turn.
                     if (currentUserTurn != null)
                     {
                         // Add the stored User turn
                         contentsList.Add(currentUserTurn);
                         // Add the current Assistant turn
                         contentsList.Add(new { role = role, parts = new[] { new { text = text } } });
                         
                         // Reset for the next pair
                         currentUserTurn = null;
                     }
                     else
                     {
                         // This Assistant turn is not preceded by a User turn. Exclude it.
                         // Optionally log: ed.WriteMessage($"\nWarning: Skipping unpaired Assistant turn: '{trimmedLine}'");
                     }
                 }
                 else
                 {
                     // Line does not start with "User:" or "Assistant:". Exclude it.
                     // If there was a pending User turn before this malformed line, it's also implicitly discarded.
                     currentUserTurn = null; // Discard any pending User turn as the sequence is broken
                                             // Optionally log: ed.WriteMessage($"\nWarning: Skipping malformed line in convo.txt: '{trimmedLine}'");
                 }
             }
             
             // After the loop, if currentUserTurn is not null, it means the last User turn
             // in the file did not have a corresponding Assistant turn. It will be excluded automatically
             // because it was never added to contentsList.
             // Optionally log: if (currentUserTurn != null) { ed.WriteMessage($"\nWarning: Last User turn in convo.txt was unpaired and excluded."); }
              
             // ed.WriteMessage($"\nLoaded conversation history from convo.txt."); // You can re-add this if you like.
         }
         else
         {
             // ed.WriteMessage($"\nWarning: conversation.txt not found at '{convofilePath}'. No conversation history loaded from file.");
         }

         //contentsList.Add(new { role = "user", parts = new[] { new { text = prompt } } });

         var lastdata = new[]
         {
             new { key = (dynamic)"user", value = (dynamic)$"{userMessage}" },
             //new { key = (dynamic)"assistant", value = (dynamic)"value2" },
             //new { key = (dynamic)"user", value = (dynamic)$"{userMessage}" },
             //new { key = (dynamic)"assistant", value = (dynamic)"value3" }
         };
          
         string filePath = Path.Combine(DataGlobal.convofilepath, "conversation.txt");
         DataGlobal.UserMessage = userMessage;

         // Data to send in the POST request

         // Combine the two arrays
         //
         var data = ((IEnumerable<dynamic>)contentsList).SelectMany(x => ((IEnumerable<dynamic>)x.parts).Select(y => new { key = x.role, value = y.text })).ToArray();
         
          var combinedData = data.ToList();

          combinedData.AddRange(lastdata);

          string jsonData = JsonConvert.SerializeObject(combinedData);
          
         if (DataGlobal.isDevMessageOn)
         {
            actDoc.Editor.WriteMessage($"\nJSON Data: {jsonData}");
         }

          // Create a new instance of HttpClient
         using HttpClient client = new HttpClient();
          // Set the content of the request
          HttpContent content = new StringContent(jsonData, Encoding.UTF8, "application/json");
          try
          {
              // Send the POST request
              HttpResponseMessage response = await client.PostAsync(url, content);
            // Check the status code of the response
            if (response.IsSuccessStatusCode)
            {
               // Get the response content
               string responseBody = response.Content.ReadAsStringAsync().Result;
               string formattedResponse = responseBody.Replace("\\n", Environment.NewLine)
                                                   .Replace("**", "")
                                                   .Replace("\\\"", "")
                                                   .Trim(' ')
                                                   .Trim('"')
                                                   .TrimStart('"')
                                                   .TrimEnd('"');
               
               if (DataGlobal.isDevMessageOn)
               {
                  actDoc.Editor.WriteMessage($"\nFormatted response: {formattedResponse}");
               }   
               string first2char = formattedResponse.ToString().Substring(0, 2);
               if (DataGlobal.isDevMessageOn)
               {
                  actDoc.Editor.WriteMessage($"\nFirst two characters: {first2char}");
               }
                  
               if (first2char == "--")
               {
                  string strcommand = formattedResponse.ToString().Remove(0, 2);

                  List<string> listParam = strcommand.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries)
                                               .Select(s => s.Trim())
                                               .ToList();

                  string methodInItem = listParam.LastOrDefault();
                  listParam.Remove(methodInItem);

                  ExecuteCommandString(listParam);

               }
               else if (first2char == ">>")
               {
                  GenaiMethods genaiMethods = new GenaiMethods();
                  string strcommand = formattedResponse.ToString().Remove(0, 2);
                  if (DataGlobal.isDevMessageOn)
                  {
                     actDoc.Editor.WriteMessage($"\nString command: {strcommand}");
                  }

                  genaiMethods.ExecuteMethod(strcommand);
               }
               else
               {
                  //content = $"assistant: {formattedResponse}";
                  // Create or append to the file
                  //File.AppendAllText(filePath, content + Environment.NewLine);

                  actDoc.Editor.WriteMessage($"\nAssistant: {formattedResponse}");
                  actDoc.Editor.WriteMessage("\n");

                  string contentTxt = $"model: {formattedResponse}";
                  // Create or append to the file
                  File.AppendAllText(filePath, contentTxt + Environment.NewLine);
                  //actDoc.Editor.WriteMessage($"\n{response.ToString()}");
               }
                  //actDoc.Editor.WriteMessage($"\n{responseBody}");

               
            }
            else
            {
               actDoc.Editor.WriteMessage("Failed to send request. Status code: " + response.StatusCode);
            }
          }
          catch (HttpRequestException ex)
          {
              actDoc.Editor.WriteMessage("An error occurred: " + ex.Message);
          }

       }


      //[CommandMethod("CAI_")]
      //public static async void AICommand()  
      private static async void AICommand()
      {
         PromptStringOptions pStrOpts = new PromptStringOptions("\nEnter instruction: ");
         pStrOpts.AllowSpaces = true;
         PromptResult pStrRes = actDoc.Editor.GetString(pStrOpts);
         if (pStrRes.Status != PromptStatus.OK) return;
         string strQuestion = pStrRes.StringResult;

         actDoc.Editor.WriteMessage($"\nYou entered: {strQuestion}");
         actDoc.Editor.WriteMessage($"\nPlease wait... Thinking...");

         var apiKey = "AIzaSyArkq8fZI22wusisF5Z3tZQ1F_o8qUkLt0"; // Your actual API Key
                                                                 // For production, consider using Environment.GetEnvironmentVariable("GEMINI_API_KEY");

         var client = new GeminiClient(apiKey);

         var prompt = strQuestion; // "What is concrete? Please answer me in one sentence.";

         var response = await client.GenerateContentAsync(prompt);
         //actDoc.Editor.WriteMessage($"\n* : {response.ToString()}");

         string filePath = Path.Combine(DataGlobal.convofilepath, "conversation.txt");
         //actDoc.Editor.WriteMessage($"\nConversation file path: {filePath}");
         string content = $"user: {prompt.ToString()}";
         // Create or append to the file
         File.AppendAllText(filePath, content + Environment.NewLine);

         string first2char = response.ToString().Substring(0, 2);
         if (first2char == "--")
         {
            string strcommand = response.ToString().Remove(0, 2);
            //actDoc.Editor.WriteMessage($"\nAI Assistant*: {strcommand.ToString()}");

            List<string> listParam = strcommand.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries)
                                         .Select(s => s.Trim())
                                         .ToList();

            string methodInItem = listParam.LastOrDefault();
            listParam.Remove(methodInItem);

            //string strresult1 = string.Join(", ", listParam);
            //actDoc.Editor.WriteMessage($"\nMethod: {methodInItem}; >{strresult1}");
            ExecuteCommandString(listParam);

            //GenaiComs genaiComs = new GenaiComs();
            //Type genaitype = genaiComs.GetType();
            //
            //MethodInfo methodInfo = genaitype.GetMethod(methodInItem);
            //
            //if (methodInfo != null)
            //{
            //    object genaicomresult = methodInfo.Invoke(genaiComs, null); // GetPoint3d takes no parameters
            //    actDoc.Editor.WriteMessage($": {genaicomresult}");
            //}
         }
         else if (first2char == ">>")
         {
            GenaiMethods genaiMethods = new GenaiMethods();
            actDoc.Editor.WriteMessage($"\n* ai string command: {response.ToString()}");
            string strcommand = response.ToString().Remove(0, 2);

            genaiMethods.ExecuteMethod(strcommand);

         }
         else
         {
            content = $"assistant: {response.ToString()}";
            // Create or append to the file
            File.AppendAllText(filePath, content + Environment.NewLine);

            actDoc.Editor.WriteMessage($"\n{content}");
            actDoc.Editor.WriteMessage("\n");
            //actDoc.Editor.WriteMessage($"\n{response.ToString()}");
         }


      }

        private static void ExecuteCommandString(List<string> parameters = null)
        {
            string commandString = "_" + parameters[0]; // Add underscore for localization
            parameters.RemoveAt(0); // Remove the command name from the parameters

            if (parameters != null && parameters.Any())
            {
                // Join parameters with spaces (mimics typing at the command line)
                // Ensure strings with spaces are enclosed in double quotes
                // For numeric inputs, just include them as is.
                commandString += " " + string.Join(" ", parameters);
            }

            commandString += "\n"; // Simulate pressing Enter to finish the command

            //actDoc.Editor.WriteMessage($"\n  [SendStringToExecute] Sending: '{commandString.TrimEnd('\n')}'");
            //actDoc.Editor.Command("AI_SelectPointM01 \n 11");
            actDoc.SendStringToExecute(commandString, true, false, true);
            //actDoc.Editor.Command(commandString);
        }
    }
}
