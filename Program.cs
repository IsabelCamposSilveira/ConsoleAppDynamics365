using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceModel.Description;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;  
using System.Net;  
using Microsoft.Xrm.Tooling.Connector;
using System.Configuration;
using Microsoft.Xrm.Sdk.Query;
using System.Windows.Controls;
using Microsoft.Crm.Sdk.Messages;

namespace ConsoleAppDynamics
{
    internal class Program
    {
        static void Main(string[] args)
        {
            // Criando a conexão
            Console.WriteLine("Conectando ao CRM. Aguarde um instante...\n");
            CRM_Connection connection = new CRM_Connection();
            CrmServiceClient connectionProd = connection.ConectarCRM_Prod();

            // Chamar as fuñções necessárias
            RetrieveContact(connectionProd);


        }

        public static void RetrieveContact(CrmServiceClient service)
        {

            // Buscar contatos 
            var fetchContact = new FetchExpression(@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='true'>
                                                        <entity name='contact'>
                                                        <attribute name='firstname'/>
                                                        <attribute name='lastname'/>                                          
                                                        </entity>
                                                    </fetch>");

            var RetrieveContact = service.RetrieveMultiple(fetchContact);
            
        }


    }
}
