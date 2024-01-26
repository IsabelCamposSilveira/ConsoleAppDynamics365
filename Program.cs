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
using Newtonsoft.Json.Linq;
using System.IdentityModel.Metadata;
using System.Web.UI.WebControls;
using System.Windows;
using System.Xml.Linq;
using Exclusao_de_Tarefas_Duplicadas.EarlyBound;

namespace ConsoleAppDynamics
{
    internal class Program
    {
        static void Main(string[] args)
        {
            // Criando a conexão
            Console.WriteLine("Conectando ao CRM. Aguarde um instante...\n");
            CRM_Connection connection = new CRM_Connection();
            CrmServiceClient connectionDev = connection.ConectarCRM_DEV();
            // CrmServiceClient connectionProd = connection.ConectarCRM_PROD();

            // Chamar as funções
            UpdateEstimatedRevenueProjectTaskEventual(connectionDev);


        }

        public static void UpdateEstimatedRevenueProjectTaskEventual(CrmServiceClient service)
        {
            var semLinhaDaOrdem = "Tarefas sem linha da ordem relacionada: \n";

            // Buscar ProjectTask do tipo Eventual
            var fetchProjectTask = new FetchExpression(@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='true'>
                                                          <entity name='msdyn_projecttask'>
                                                            <attribute name='msdyn_subject'/>
                                                            <attribute name='msdyn_projecttaskid'/>
                                                            <attribute name='smt_horasestimadas'/>
                                                            <attribute name='msdyn_project'/>
                                                            <attribute name='msdyn_remaininghours'/>
                                                            <attribute name='msdyn_effort'/>
                                                            <attribute name='smt_lp_milestone'/>
                                                            <order attribute='msdyn_subject' descending='false'/>
                                                            <filter type='and'>
                                                              <condition attribute = 'smt_lp_milestone' operator= 'not-null'/> 
                                                              <condition attribute = 'smt_mn_estimated_revenue' operator= 'null'/>
                                                              <condition attribute = 'msdyn_project' operator= 'not-null'/>
                                                            </filter >
                                                             <link-entity name='smt_milestone_project' from='smt_milestone_projectid' to='smt_lp_milestone' link-type='inner' alias='ab'>
                                                              <filter type='and'>
                                                                <condition attribute='smt_pl_type' operator='eq' value='1'/>
                                                              </filter>
                                                             </link-entity>                                              
                                                          </entity>
                                                        </fetch>");
            var RetrieveProjectTask = service.RetrieveMultiple(fetchProjectTask);

            // Para cada ProjectTask encontrada
            foreach (var ProjectTask in RetrieveProjectTask.Entities) {

                // Buscar Linha da ordem relacionada 
                var project = ProjectTask.GetAttributeValue<EntityReference>("msdyn_project");
                var fetchlinhaDaOrdem = new FetchExpression(@"<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                      <entity name='salesorderdetail'>
                        <attribute name='salesorderdetailid'/>
                        <attribute name='smt_mn_valor_liq'/>
                        <attribute name='priceperunit'/>
                        <order attribute='priceperunit' descending='false'/>
                        <filter type='and'>
                          <condition attribute='msdyn_project' operator='eq' uitype='msdyn_project' value='{" + project.Id + @"}'/>
                          <condition attribute='producttypecode' operator='eq' value='5'/>
                        </filter>
                      </entity>
                    </fetch>");
                var RetrievelinhaDaOrdem = service.RetrieveMultiple(fetchlinhaDaOrdem);

                // Se encontrar uma linha relacionada 
                if (RetrievelinhaDaOrdem.Entities.Count > 0)
                {
                    // Buscar Marco
                    var MarcoEntity = ProjectTask.GetAttributeValue<EntityReference>("smt_lp_milestone");
                    var Marco = service.Retrieve("smt_milestone_project", MarcoEntity.Id, new ColumnSet("smt_dc_stage"));

                    // Utiliza a primeira linha da ordem encontrada
                    var linhaDaOrdem = RetrievelinhaDaOrdem.Entities.FirstOrDefault();

                    // var usadas para o calculo
                    var Moneysmt_mn_valor_liq = (Money)linhaDaOrdem["smt_mn_valor_liq"];
                    var smt_mn_valor_liq = (Decimal)Moneysmt_mn_valor_liq.Value;
                    var smt_dc_stage = (Decimal)Marco["smt_dc_stage"];
                    var MoneyPricePerUnit = (Money)linhaDaOrdem["priceperunit"];
                    var PricePerUnit = (Decimal)MoneyPricePerUnit.Value;

                    // calculos para o Revenue
                    decimal? estimatedRevenue = linhaDaOrdem["priceperunit"] != null ? (smt_dc_stage * PricePerUnit) / 100 : 0;
                    decimal? liquidEstimated = linhaDaOrdem["smt_mn_valor_liq"] != null ? (smt_dc_stage * smt_mn_valor_liq) / 100 : 0;

                    // Atualizar a entidade msdyn_projecttask
                    msdyn_projecttask msdyn_projecttask = new msdyn_projecttask();
                    msdyn_projecttask.Id = ProjectTask.Id;
                    msdyn_projecttask.smt_mn_estimated_revenue = new Money(estimatedRevenue.Value);
                    msdyn_projecttask.smt_mn_estimated_liquidrevenue = new Money(liquidEstimated.Value);
                    service.Update(msdyn_projecttask);
                }               
                else
                {
                    // Se não encontrar uma linha relacionada saltar o id
                    semLinhaDaOrdem = semLinhaDaOrdem + ProjectTask.Id + " ;\n";
                }
            }

            Console.WriteLine(semLinhaDaOrdem);
        }


    }
}
