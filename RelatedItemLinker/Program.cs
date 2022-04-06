using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security;
using System.Windows.Forms;

namespace RelatedItemLinker
{
    internal static class Program
    {
        static void Main(string[] args)
        {
            if (args.Length == 0)
            {
                showError("Faltam argumentos");
                showHelp();
                return;
            }
            if (args.Length != 10)
            {
                showError("Quantidade de argumentos invalida");
                showHelp();
                return;
            }

            try
            {
                Linker linker = getLinker(args);

                //Console.Write("Informe o usuário de acesso ao Sharepoint Online: ");
                //string username = Console.ReadLine();
                //Console.Write("\r\nSenha: ");
                //SecureString pwd = getPassword();
                //SharePointOnlineCredentials creds = new SharePointOnlineCredentials(username, pwd);
                //ClientContext cli = new ClientContext(linker.siteUrl);
                //cli.Credentials = creds;

                var authManager = new OfficeDevPnP.Core.AuthenticationManager();
                ClientContext cli = authManager.GetWebLoginClientContext(linker.siteUrl);

                Write("Validando lista de origem... ", ConsoleColor.White);
                List lst = cli.Web.Lists.GetById(linker.listID);
                cli.Load(lst);
                cli.Load(cli.Web);  
                cli.ExecuteQuery();
                WriteLine("Ok", ConsoleColor.Green);
                Guid webId = cli.Web.Id;

                Write("Validando item de origem... ", ConsoleColor.White);
                ListItem item = lst.GetItemById(linker.itemID);
                cli.Load(item);
                cli.ExecuteQuery();
                WriteLine("Ok", ConsoleColor.Green);

                Write("Validando lista de tarefas de destino... ", ConsoleColor.White);
                List tlst = cli.Web.Lists.GetById(linker.taskListID);
                cli.Load(tlst);
                cli.ExecuteQuery();
                WriteLine("Ok", ConsoleColor.Green);

                Write("Validando tarefa de destino... ", ConsoleColor.White);
                ListItem titem = tlst.GetItemById(linker.taskItemID);
                cli.Load(titem);
                cli.ExecuteQuery();
                WriteLine("Ok", ConsoleColor.Green);

                WriteLine("Alterando campo RelatedItems...", ConsoleColor.White);
                titem["RelatedItems"] = string.Format("[{{\"ItemId\":{0},\"WebId\":\"{1}\",\"ListId\":\"{2}\"}}]",
                    linker.itemID, webId.ToString().Replace("{", string.Empty).Replace("}", string.Empty),
                    linker.listID.ToString().Replace("{", string.Empty).Replace("}", string.Empty));
                WriteLine(titem["RelatedItems"].ToString(), ConsoleColor.Blue);
                //titem.Update();
                //cli.ExecuteQuery();
                WriteLine("Alteração efetuada com sucesso!", ConsoleColor.Green);

                //[{"ItemId":<id-do-item>,"WebId":"<guid-do-site>","ListId":"<guid-da-lista>"}]
            }
            catch (Exception ex)
            {
                Console.WriteLine("");
                showError(ex.Message);
            }
            Console.ReadLine();


        }

        static void showHelp()
        {
            Console.WriteLine(@"Sintaxe:
	relateditemlinker -url <site> -list <guid> -item <int> -tasklist <guid> -taskitem <int>

Onde:
	-url <site> - URL do site onde os itens estao localizados
	-list <guid> - Identificador no formato GUID da lista de origem do item a ser referenciado
	-item <int> - Identificador inteiro do item a ser referenciado
	-tasklist <guid> - Identificador no formato GUID da lista de tarefas
	-taskitem <int> - Identificador inteiro da tarefa onde o item será referenciado");
        }

        static void WriteLine(string message, ConsoleColor color)
        {
            Console.ForegroundColor = color;
            Console.WriteLine(message);
            Console.ResetColor();
        }
        static void Write(string message, ConsoleColor color)
        {
            Console.ForegroundColor = color;
            Console.Write(message);
            Console.ResetColor();
        }

        static SecureString getPassword()
        {
            var pwd = new SecureString();
            while (true)
            {
                ConsoleKeyInfo i = Console.ReadKey(true);
                if (i.Key == ConsoleKey.Enter)
                {
                    break;
                }
                else if (i.Key == ConsoleKey.Backspace)
                {
                    if (pwd.Length > 0)
                    {
                        pwd.RemoveAt(pwd.Length - 1);
                        Console.Write("\b \b");
                    }
                }
                else if (i.KeyChar != '\u0000') // KeyChar == '\u0000' if the key pressed does not correspond to a printable character, e.g. F1, Pause-Break, etc
                {
                    pwd.AppendChar(i.KeyChar);
                    Console.Write("*");
                }
            }
            return pwd;
        }

        static Linker getLinker(string[] args)
        {
            Linker l = new Linker();
            for(var ndx = 0; ndx < args.Length; ndx++)  
            {
                string arg = args[ndx];
                if (arg == "-list")
                    l.listID = Guid.Parse(args[ndx + 1]);

                else if (arg == "-item")
                    l.itemID = int.Parse(args[ndx + 1]);

                else if (arg == "-tasklist")
                    l.taskListID = Guid.Parse(args[ndx + 1]);

                else if (arg == "-taskitem")
                    l.taskItemID = int.Parse(args[ndx + 1]);

                else if (arg == "-site")
                    l.siteUrl = args[ndx + 1];
            }
            return l;
        }
        static void showError(string msg)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine(msg);
            Console.ResetColor();
        }
    }

    public class Linker
    {
        public Guid listID { get; set; }
        public int itemID { get; set; }
        public Guid taskListID { get; set; }
        public int taskItemID { get; set; }
        public string siteUrl { get; set; }
    }
}
