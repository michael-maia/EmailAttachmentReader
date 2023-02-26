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
            // Variáveis ​​que conterão o valor do intervalo de tempo que está escrito no arquivo config.ini
            int timeIntervalHours = 0;
            int timeIntervalMinutes = 0;

            // Verificando a data em que o programa está sendo executado para que possamos transformá-lo em uma string e salvá-la para uso no nome do arquivo de log
            string actualDate = DateTime.Now.ToString("dd-MM-yyyy");

            // Verificando se o arquivo config.ini existe (arquivo de configuração)
            if (File.Exists("config.ini") == true)
            {
                try
                {
                    // Se existir ele irá ler os valores por intervalo de tempo para que a configuração possa ser atualizada
                    using (StreamReader sr = new StreamReader("config.ini"))
                    {
                        string line;

                        while ((line = sr.ReadLine()) != null)
                        {
                            if (line.StartsWith("HORAS"))
                            {
                                string[] lineSplit = line.Trim().Split('=');
                                timeIntervalHours = int.Parse(lineSplit[1]);
                            }
                            else if (line.StartsWith("MINUTOS"))
                            {
                                string[] lineSplit = line.Trim().Split('=');
                                timeIntervalMinutes = int.Parse(lineSplit[1]);
                            }
                        }
                    }
                }
                catch (FormatException e)
                {
                    using (StreamWriter sw = new StreamWriter($"logs\\log_{actualDate}.txt", append: true))
                    {
                        Console.WriteLine($"[ERROR: {DateTime.Now}] Verifique os valores de entrada no config.ini\nMessagem => {e.Message}");
                        sw.WriteLine($"[ERROR: {DateTime.Now}] Verifique os valores de entrada no config.ini\nMessagem => {e.Message}");
                        sw.Close();
                    }
                    PressKeyToContinue();
                    
                    // Encerra aplicação
                    Environment.Exit(1);
                }
            }
            else
            {
                // Caso não exista será criado um arquivo padrão com 1 hora e 0 minutos como valores
                using (StreamWriter sw = new StreamWriter("config.ini"))
                {
                    sw.WriteLine("[INTERVALO_TEMPO]");
                    sw.WriteLine("HORAS=1");
                    sw.WriteLine("MINUTOS=0");
                    sw.Close();
                }

                // Valores padrão quando o arquivo é criado
                timeIntervalHours = 1;
                timeIntervalMinutes = 0;

                using (StreamWriter sw = new StreamWriter($"logs\\log_{actualDate}.txt", append: true))
                {
                    Console.WriteLine($"[{DateTime.Now}] O arquivo config.ini foi criado");
                    sw.WriteLine($"[{DateTime.Now}] O arquivo config.ini foi criado");
                    sw.Close();
                }
            }

            // Repetição infinita pois o email que estamos procurando pode ver a qualquer momento do dia
            while (true)
            {                
                // Adicionando User Secrets para ler o que foi guardado
                var config = new ConfigurationBuilder().AddUserSecrets<Program>().Build();
                
                // Verificando se a pasta 'logs' existe, pois é onde o programa vai guardar todos dados de log
                if (!Directory.Exists("logs"))
                {
                    Directory.CreateDirectory("logs");
                }               
                
                // Iniciando a transmissão de dados para o arquivo de log, mas toda vez que o programa for executado ele gravará dentro do mesmo log daquele dia
                using (StreamWriter sw = new StreamWriter($"logs\\log_{actualDate}.txt", append: true))
                {
                    try
                    {                        
                        // Criando um cliente OpenPop para que possamos acessar nossos e-mails
                        using (Pop3Client client = new())
                        {                            
                            // Conectando no servidor de email
                            client.Connect(config["AuthenticationData:Hostname"], int.Parse(config["AuthenticationData:Port"]), bool.Parse(config["AuthenticationData:UseSSL"]));
                            Console.WriteLine($"[{DateTime.Now}] Conectado no servidor de email");
                            sw.WriteLine($"[{DateTime.Now}] Conectado no servidor de email");
                            
                            // Usando nosso cliente para autenticar no servidor
                            client.Authenticate(config["AuthenticationData:Email"], config["AuthenticationData:Password"], AuthenticationMethod.UsernameAndPassword);
                            Console.WriteLine($"[{DateTime.Now}] Cliente autenticado no servidor");
                            sw.WriteLine($"[{DateTime.Now}] Cliente autenticado no servidor");
                            
                            // Verificando o número de mensagens na caixa de entrada
                            int messageCount = client.GetMessageCount();
                            Console.WriteLine($"[{DateTime.Now}] Número de e-mails: {messageCount}");
                            sw.WriteLine($"[{DateTime.Now}] Número de e-mails: {messageCount}");
                            
                            // Só precisamos fazer todo o código dentro deste escopo se houver pelo menos 1 e-mail na caixa de entrada
                            if (messageCount > 0)
                            {                                
                                // Esta lista irá armazenar todos os e-mails que recebemos
                                List<Message> allMessages = new(messageCount);

                                // As mensagens são numeradas no intervalo: [1, messageCount]
                                // Logo, os números das mensagens são baseados em 1
                                // Most servers give the latest message the highest number
                                for (int i = messageCount; i > 0; i--)
                                {
                                    allMessages.Add(client.GetMessage(i));
                                }
                                
                                // Caminho para onde os anexos serão transferidos
                                string targetPath = config["Others:TargetPath"];

                                // Se a pasta de destino não existir, ele criará uma para que o programa continue rodando sem erros
                                if (!Directory.Exists(targetPath))
                                {
                                    Directory.CreateDirectory(targetPath);
                                    Console.WriteLine($"[{DateTime.Now}] A pasta {targetPath} foi criada!");
                                    sw.WriteLine($"[{DateTime.Now}] A pasta {targetPath} foi criada!");                                    
                                }
                                
                                // Lendo cada email recebido
                                foreach (Message message in allMessages)
                                {
                                    // Apenas verificando se é de um email específico
                                    if (message.Headers.From.Address.Trim() == config["EmailReceived:Address"])
                                    {                                        
                                        // Saving all attachments in a list so it will be checked one by one
                                        List<MessagePart> attachments = message.FindAllAttachments();
                                        foreach (var attachment in attachments)
                                        {
                                            // Salvando todos os anexos em uma lista para que sejam verificados um por um
                                            if (attachment.FileName.StartsWith(config["EmailReceived:Attachment1"]) || attachment.FileName.StartsWith(config["EmailReceived:Attachment2"]))
                                            {
                                                Console.WriteLine($"[{DateTime.Now}] Transferindo {attachment.FileName} para a pasta alvo");
                                                sw.WriteLine($"[{DateTime.Now}] Transferindo {attachment.FileName} para a pasta alvo");

                                                // Todos os anexos da mensagem serão salvos no targetPath
                                                File.WriteAllBytes(Path.Combine(targetPath, attachment.FileName), attachment.Body);
                                            }
                                        }
                                    }
                                }

                                // Depois de salvar todos os anexos de cada mensagem, os emails serão excluídos, pois só precisamos verificar os recentes
                                client.DeleteAllMessages();
                                Console.WriteLine($"[{DateTime.Now}] Email foi removido da caixa de entrada");
                                sw.WriteLine($"[{DateTime.Now}] Email foi removido da caixa de entrada");
                            }
                            // Quando não há email na caixa de entrada
                            else
                            {
                                Console.WriteLine($"[{DateTime.Now}] Não há email na caixa de entrada");
                                sw.WriteLine($"[{DateTime.Now}] Não há email na caixa de entrada");
                            }

                            // Antes de terminar o processo, ele se desconectará automaticamente do servidor de email
                            client.Dispose();
                            Console.WriteLine($"[{DateTime.Now}] Desconectando do servidor de email");
                            sw.WriteLine($"[{DateTime.Now}] Desconectando do servidor de email");
                        }

                        // Fechando a transmissão de dados no log para que o arquivo seja salvo
                        sw.Close();

                        // Converting time in config.ini file to miliseconds
                        int hourInMiliseconds = timeIntervalHours * 3600000;
                        int minuteInMiliseconds = timeIntervalMinutes * 1000;

                        // Coloca o programa em 'hibernação' por 1 hora e depois roda novamente
                        Console.WriteLine($"\n\nProcesso em stand-by, próxima vez que vai rodar será no {DateTime.Now.Add(new TimeSpan(0, timeIntervalHours, timeIntervalMinutes, 0))}\n\n");
                        Thread.Sleep(hourInMiliseconds + minuteInMiliseconds);
                        Console.Clear();
                    }
                    // Algumas exceções que podem acontecer durante a execução
                    catch (PopServerNotFoundException e)
                    {
                        Console.WriteLine($"[ERROR: {DateTime.Now}] Conexão com o servidor de email não é possível!\nMessagem => {e.Message}");
                        sw.WriteLine($"[ERROR: {DateTime.Now}] Conexão com o servidor de email não é possível!\nMessagem => {e.Message}");
                        PressKeyToContinue();
                        break;
                    }
                    catch (InvalidLoginException e)
                    {
                        Console.WriteLine($"[ERROR: {DateTime.Now}] Usuário e/ou senha inválidos ao autenticar no servidor\nMessagem => {e.Message}");
                        sw.WriteLine($"[ERROR: {DateTime.Now}] Usuário e/ou senha inválidos ao autenticar no servidor\nMessagem => {e.Message}");
                        PressKeyToContinue();
                        break;
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine($"[ERROR: {DateTime.Now}] Error! Verifique a mensagem abaixo\nMessagem => {e.Message}");
                        sw.WriteLine($"[ERROR: {DateTime.Now}] Error! Verifique a mensagem abaixo\nMessagem => {e.Message}");
                        PressKeyToContinue();
                        break;
                    }
                }
            }
        }
        private static void PressKeyToContinue()
        {
            Console.WriteLine("Pressione um botão para sair da aplicação...");
            Console.ReadKey();
        }
    }
}