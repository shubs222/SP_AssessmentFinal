using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Security;

namespace SPAssessment
{
    class UserCredentials
    {
        public static string UserName;
        public static SecureString Passwrd;
       
        public static void Getdata()
        {
          
                Console.WriteLine("Enter user name");
                UserName = Console.ReadLine();
                Console.WriteLine("Enter password");
                Passwrd = GetPassword();
          
        }
        private static SecureString GetPassword()
        {
            ConsoleKeyInfo info;
            //Get the user's password as a SecureString  
            SecureString securePassword = new SecureString();
            do
            {
                info = Console.ReadKey(true);
                if (info.Key != ConsoleKey.Enter)
                {
                    securePassword.AppendChar(info.KeyChar);
                }
            }
            while (info.Key != ConsoleKey.Enter);
            return securePassword;
        }


    }
}
