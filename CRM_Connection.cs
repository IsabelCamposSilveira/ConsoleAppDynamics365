using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceModel.Description;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Xrm.Sdk;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Client;  // Adicionado
using System.Net;  // Adicionado
using Microsoft.Xrm.Tooling.Connector;
using System.Configuration;
using Microsoft.Xrm.Sdk.Query;
using System.Windows.Controls;


namespace ConsoleAppDynamics
{
    public class CRM_Connection
    {
        public CrmServiceClient ConectarCRM_Prod()
        {
            try
            {
                // Dados da conexão
                string url = ""; // Inserir URL do ambiente
                string userName = ""; // Inserir e-mail do login
                string password = ""; // Inserir senha

                string connectionString = $"AuthType=Office365;Url={url};Username={userName};Password={password}";

                // Criando a conexão
                CrmServiceClient service = new CrmServiceClient(connectionString);

                return service;
            }             
            catch (Exception ex)
            {
                throw new ArgumentNullException("Falha no Método [ConectarCRM]: " + ex.Message);
            }         

        }


    }
} 
