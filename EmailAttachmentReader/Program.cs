using Microsoft.Extensions.Configuration;
using OpenPop.Mime;
using OpenPop.Pop3;
using OpenPop.Pop3.Exceptions;

namespace EmailAttachmentReader
{
    internal class Program
    {
        static void Main()
        {
            // Variables that will hold the value of time interval that is written on the config.ini file
            int timeIntervalHours = 0;
            int timeIntervalMinutes = 0;

            // Checking the date when the program is running so we can turn into a string and save it for use in the log filename
            string actualDate = DateTime.Now.ToString("dd-MM-yyyy");

            // Checking if the folder 'logs' exists, because is where the program will store all log data
            if (!Directory.Exists("logs"))
            {
                Directory.CreateDirectory("logs");
            }

            // Checking if config.ini file exists (configuration file)
            if (File.Exists("config.ini") == true)
            {
                try
                {
                    // If exists it will read the values for time interval so the configuration can be updated
                    using (StreamReader sr = new StreamReader("config.ini"))
                    {
                        string line;

                        while ((line = sr.ReadLine()) != null)
                        {
                            if (line.StartsWith("HOURS"))
                            {
                                string[] lineSplit = line.Trim().Split('=');
                                timeIntervalHours = int.Parse(lineSplit[1]);
                            }
                            else if (line.StartsWith("MINUTES"))
                            {
                                string[] lineSplit = line.Trim().Split('=');
                                timeIntervalMinutes = int.Parse(lineSplit[1]);
                            }
                        }
                        sr.Close();
                    }

                    if (timeIntervalHours < 0 || timeIntervalMinutes < 0)
                    {
                        using (StreamWriter sw = new StreamWriter($"logs\\log_{actualDate}.txt", append: true))
                        {
                            Console.WriteLine($"[{DateTime.Now}] The time interval value cannot be negative! Change it on config.ini");
                            sw.WriteLine($"[{DateTime.Now}] The time interval value cannot be negative! Change it on config.ini");
                            sw.Close();
                        }

                        PressKeyToContinue();
                        Environment.Exit(1);
                    }
                }
                catch(FormatException e)
                {
                    using (StreamWriter sw = new StreamWriter($"logs\\log_{actualDate}.txt", append: true))
                    {
                        Console.WriteLine($"[ERROR: {DateTime.Now}] Check input values on config.ini\nMessage => {e.Message}");
                        sw.WriteLine($"[ERROR: {DateTime.Now}] Check input values on config.ini\nMessage => {e.Message}");
                        sw.Close();                        
                    }

                    PressKeyToContinue();                    
                    Environment.Exit(1);
                }
            }
            else
            {
                // If dont exists it will be created a default file with 1 hour and 0 minutes as values
                using (StreamWriter sw = new StreamWriter("config.ini"))
                {
                    sw.WriteLine("[TIME_INTERVAL]");
                    sw.WriteLine("HOURS=1");
                    sw.WriteLine("MINUTES=0");
                    sw.Close();
                }

                // Default values when file is created
                timeIntervalHours = 1;
                timeIntervalMinutes = 0;

                using (StreamWriter sw = new StreamWriter($"logs\\log_{actualDate}.txt", append: true))
                {
                    Console.WriteLine($"[{DateTime.Now}] The file config.ini was created");
                    sw.WriteLine($"[{DateTime.Now}] The file config.ini was created");
                    sw.Close();
                }
            }

            // Infinite loop because the email we are looking for can come anytime of the day
            while (true)
            {
                // Updating the date for creating a log file per day
                actualDate = DateTime.Now.ToString("dd-MM-yyyy");

                // Adding User Secrets to read what it's stored
                var config = new ConfigurationBuilder().AddUserSecrets<Program>().Build();                              

                // Starting data transmission to the log file, but everytime the program runs it will write inside the same log of that day
                using (StreamWriter sw = new StreamWriter($"logs\\log_{actualDate}.txt", append: true))
                {
                    try
                    {
                        // Creating a OpenPop client so we can access our emails
                        using (Pop3Client client = new())
                        {
                            // Connecting to e-mail server
                            client.Connect(config["AuthenticationData:Hostname"], int.Parse(config["AuthenticationData:Port"]), bool.Parse(config["AuthenticationData:UseSSL"]));
                            Console.WriteLine($"[{DateTime.Now}] Connected to the e-mail server");
                            sw.WriteLine($"[{DateTime.Now}] Connected to the e-mail server");

                            // Using our client to autenticate on the server
                            client.Authenticate(config["AuthenticationData:Email"], config["AuthenticationData:Password"], AuthenticationMethod.UsernameAndPassword);
                            Console.WriteLine($"[{DateTime.Now}] Client authenticated on the server");
                            sw.WriteLine($"[{DateTime.Now}] Client authenticated on the server");

                            // Checking the number of messages on inbox
                            int messageCount = client.GetMessageCount();
                            Console.WriteLine($"[{DateTime.Now}] Number of e-mails: {messageCount}");
                            sw.WriteLine($"[{DateTime.Now}] Number of e-mails: {messageCount}");
                            
                            // We only need to do all the code inside this scope if there is at least 1 e-mail on inbox
                            if (messageCount > 0)
                            {
                                // This list will store all the e-mails received
                                List<Message> allMessages = new(messageCount);

                                // Messages are numbered in the interval: [1, messageCount]
                                // Ergo: message numbers are 1-based.
                                // Most servers give the latest message the highest number
                                for (int i = messageCount; i > 0; i--)
                                {
                                    allMessages.Add(client.GetMessage(i));
                                }

                                // Path where the attachments will be transfered
                                string targetPath = config["Others:TargetPath"];
                                
                                // If the target folder don't exist, it will create one so the program will keep running without errors
                                if (!Directory.Exists(targetPath))
                                {
                                    Directory.CreateDirectory(targetPath);
                                    Console.WriteLine($"[{DateTime.Now}] The folder {targetPath} was created!");
                                    sw.WriteLine($"[{DateTime.Now}] The folder {targetPath} was created!");                                    
                                }

                                // Reading every e-mail received
                                foreach (Message message in allMessages)
                                {
                                    // Just checking if its from a specific e-mail (this can be removed)
                                    if (message.Headers.From.Address.Trim() == config["EmailReceived:Address"])
                                    {                                        
                                        // Saving all attachments in a list so it will be checked one by one
                                        List<MessagePart> attachments = message.FindAllAttachments();
                                        foreach (var attachment in attachments)
                                        {
                                            // In this logic we are looking for two specific files (can be removed too)
                                            if (attachment.FileName.StartsWith(config["EmailReceived:Attachment1"]) || attachment.FileName.StartsWith(config["EmailReceived:Attachment2"]))
                                            {
                                                Console.WriteLine($"[{DateTime.Now}] Transfering {attachment.FileName} to the target folder");
                                                sw.WriteLine($"[{DateTime.Now}] Transfering {attachment.FileName} to the target folder");

                                                // All attachments of the message will be saved on the targetPath
                                                File.WriteAllBytes(Path.Combine(targetPath, attachment.FileName), attachment.Body);
                                            }
                                        }
                                    }
                                }

                                // After saving all the attachments of every message, the e-mails will be deleted because we only need to check the recent ones
                                client.DeleteAllMessages();
                                Console.WriteLine($"[{DateTime.Now}] E-mail were removed from the inbox");
                                sw.WriteLine($"[{DateTime.Now}] E-mail were removed from the inbox");
                            }
                            // When there is no e-mail on inbox
                            else
                            {
                                Console.WriteLine($"[{DateTime.Now}] There is no e-mail on inbox");
                                sw.WriteLine($"[{DateTime.Now}] There is no e-mail on inbox");
                            }

                            // Before finishing the process, it will auto disconnect from the email server
                            client.Dispose();
                            Console.WriteLine($"[{DateTime.Now}] Disconnecting from the email server");
                            sw.WriteLine($"[{DateTime.Now}] Disconnecting from the email server");
                        }

                        // Closing the data transmission to the log so the file will be saved
                        sw.Close();                                                 

                        // Converting time in config.ini file to miliseconds
                        int hourInMiliseconds = timeIntervalHours * 3600000;
                        int minuteInMiliseconds = timeIntervalMinutes * 1000;

                        // Puts the program to 'sleep' for 1 hour and then it will run again
                        Console.WriteLine($"\n\nProcess in stand-by, next run will be at {DateTime.Now.Add(new TimeSpan(0, timeIntervalHours, timeIntervalMinutes, 0))}\n\n");
                        Thread.Sleep(hourInMiliseconds + minuteInMiliseconds);                        
                        Console.Clear();
                    }
                    // Some exceptions that can happen while it's running
                    catch (PopServerNotFoundException e)
                    {
                        Console.WriteLine($"[ERROR: {DateTime.Now}] Connection to the e-mail server is not possible!\nMessage => {e.Message}");
                        sw.WriteLine($"[ERROR: {DateTime.Now}] Connection to the e-mail server is not possible!\nMessage => {e.Message}");
                        PressKeyToContinue();
                        break;
                    }
                    catch (InvalidLoginException e)
                    {
                        Console.WriteLine($"[ERROR: {DateTime.Now}] Invalid user and/or password when authenticating on the server\nMessage => {e.Message}");
                        sw.WriteLine($"[ERROR: {DateTime.Now}] Invalid user and/or password when authenticating on the server\nMessage => {e.Message}");
                        PressKeyToContinue();
                        break;
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine($"[ERROR: {DateTime.Now}] Error! Check the message below\nMessage => {e.Message}");
                        sw.WriteLine($"[ERROR: {DateTime.Now}] Error! Check the message below\nMessage => {e.Message}");
                        PressKeyToContinue();
                        break;
                    }
                }
            }
        }
        private static void PressKeyToContinue()
        {
            Console.WriteLine("Press a key to exit the application...");
            Console.ReadKey();
        }
    }
}