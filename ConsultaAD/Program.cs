using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.DirectoryServices;
using System.IO;

namespace ConsultaAD
{
    class Program
    {
        static void Main(string[] args)
        {          
            var _directoryEntry = new DirectoryEntry("OU a donde conectarse", "Usuario", "Clave");

            var user = new List<string>();
            foreach (DirectoryEntry child in _directoryEntry.Children)
            {
                if (child.Name.Contains("OU="))   
                    user.AddRange(GetUserAD(child.Properties["distinguishedName"].Value.ToString()));
                else
                {
                    var IsPersona = false;
                    object values = child.Properties["objectClass"].Value;
                    Array arObjectClass = (Array)values;
                    foreach (var item in arObjectClass)
                    {
                        if (item.ToString().Contains("person"))
                        {
                            IsPersona = true;
                            break;
                        }
                    }
                    if (!IsPersona)
                    {
                        user.AddRange(GetUserAD(child.Properties["distinguishedName"].Value.ToString()));
                    }
                    else if (!string.IsNullOrEmpty((string)child.Properties["mail"].Value))
                        user.Add(child.Name + ";" + (string)child.Properties["sAMAccountName"].Value + ";" + (string)child.Properties["mail"].Value);
                }         
            }

            var path = "usuariosAD"+DateTime.Now.ToShortDateString()+".csv";
            MemoryStream stream = new MemoryStream();
            if (!File.Exists(path))
            {
                var file = File.CreateText(path);
                file.Close();
            }
            byte[] buffer = File.ReadAllBytes(path);
            byte[] newData = System.Text.Encoding.UTF8.GetBytes(string.Join(Environment.NewLine, user));
            stream.Write(newData, 0, newData.Length);
            stream.Write(buffer, 0, buffer.Length);
            File.WriteAllBytes(path, stream.GetBuffer());
            stream.Dispose();
        }    

        private static List<string> GetUserAD(string OUCNData)
        {
            Console.WriteLine("Evaluando: " + OUCNData);
            if (!OUCNData.ToLower().Contains("user") && !OUCNData.ToLower().Contains("usuar")) return user;
            var user = new List<string>();
            var _directoryEntryOU = new DirectoryEntry("LDAP Path" + OUCNData, "USUARIO", "CLAVE");
            Console.WriteLine("Procesando: " + OUCNData);
            foreach (DirectoryEntry child in _directoryEntryOU.Children)
            {
                var IsPersona = false;
                object values = child.Properties["objectClass"].Value;
                Array arObjectClass = (Array)values;
                foreach (var item in arObjectClass)
                {
                    if (item.ToString().Contains("person"))
                    {
                        IsPersona = true;
                        break;
                    }
                }
                if (!IsPersona)
                    user.AddRange(GetUserAD(child.Properties["distinguishedName"].Value.ToString()));
                else 
                    user.Add(child.Name + ";" + (string)child.Properties["sAMAccountName"].Value + ";" + (string)child.Properties["mail"].Value);
            }
            return user;
        }
    }
}
