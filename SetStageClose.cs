using System;
using System.Activities;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.Remoting.Contexts;
using System.ServiceModel;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Workflow;

namespace ActiveStageChangetoClosed
{

    public class SetStageClose : CodeActivity
    {
        [Input("caseid")]
        [ReferenceTarget("tc_case")]
        [RequiredArgument]

        public InArgument<EntityReference> caseid { get; set; }

        [Input("category")]
        [RequiredArgument]
        public InArgument<string> category { get; set; }

        protected override void Execute(CodeActivityContext context)
        {

            Guid processid = Guid.Empty;
            Guid stageid = Guid.Empty;
            Guid bpfid = Guid.Empty;
            IWorkflowContext workflowContext = context.GetExtension<IWorkflowContext>();
            IOrganizationServiceFactory serviceFactory = context.GetExtension<IOrganizationServiceFactory>();

            // Use the context service to create an instance of IOrganizationService.             
            IOrganizationService service = serviceFactory.CreateOrganizationService(workflowContext.InitiatingUserId);
            Guid id = this.caseid.Get(context).Id;
            string category = this.category.Get(context);

            //Get Process ID for BPF "Case Progress Stage BPF"
            EntityCollection workflowRecords = new EntityCollection();
            QueryExpression query = new QueryExpression("workflow");
            query.NoLock = true;
            query.ColumnSet = new ColumnSet("workflowid");
            query.Criteria.AddCondition("name", ConditionOperator.Equal, "Case Progress Stage");
            workflowRecords = service.RetrieveMultiple(query);
            if (workflowRecords.Entities.Count > 0)
            {
                processid = workflowRecords.Entities.FirstOrDefault().GetAttributeValue<Guid>("workflowid");
            }
            if (category == "investigation")
            {
                //Get Stage ID for Stage "Closedinvestigation"
                EntityCollection stageRecords = new EntityCollection();
                QueryExpression query2 = new QueryExpression("processstage");
                query2.NoLock = true;
                query2.ColumnSet = new ColumnSet("processstageid");
                query2.Criteria.AddCondition("stagename", ConditionOperator.Equal, "Investigationclosed");
                query2.Criteria.AddCondition("processid", ConditionOperator.Equal, processid);
                stageRecords = service.RetrieveMultiple(query2);
                if (stageRecords.Entities.Count > 0)
                {
                    stageid = stageRecords.Entities.FirstOrDefault().GetAttributeValue<Guid>("processstageid");
                }
            }

            else if (category == "latesubmission")

            {
                //Get Stage ID for Stage "Closedinvestigation"
                EntityCollection stageRecords = new EntityCollection();
                QueryExpression query2 = new QueryExpression("processstage");
                query2.NoLock = true;
                query2.ColumnSet = new ColumnSet("processstageid");
                query2.Criteria.AddCondition("stagename", ConditionOperator.Equal, "Latesubmissionclosed");
                query2.Criteria.AddCondition("processid", ConditionOperator.Equal, processid);
                stageRecords = service.RetrieveMultiple(query2);
                if (stageRecords.Entities.Count > 0)
                {
                    stageid = stageRecords.Entities.FirstOrDefault().GetAttributeValue<Guid>("processstageid");
                }

            }

            //Get Business Process Flow  ID for the case ID 
            EntityCollection bpfRecords = new EntityCollection();
            QueryExpression querycasebpf = new QueryExpression("tc_casesprogressstage");
            querycasebpf.NoLock = true;
            querycasebpf.ColumnSet = new ColumnSet("businessprocessflowinstanceid");
            querycasebpf.Criteria.AddCondition("bpf_tc_caseid", ConditionOperator.Equal, id);
            bpfRecords = service.RetrieveMultiple(querycasebpf);
            if (bpfRecords.Entities.Count > 0)
            {

                bpfid = bpfRecords.Entities.FirstOrDefault().GetAttributeValue<Guid>("businessprocessflowinstanceid");
                //Change the stage
                Entity updatedStageBPF = new Entity("tc_casesprogressstage"); //Parent entity on which the BPF is
                updatedStageBPF.Id = bpfid;
                updatedStageBPF["activestageid"] = new EntityReference("processstage", stageid);
                service.Update(updatedStageBPF);
                //Change the stage
                Entity updatedStageCase = new Entity("tc_case"); //Parent entity on which the BPF is
                updatedStageCase.Id = id;

                updatedStageCase["stageid"] = stageid;
                //updatedStage["processid"] = processid;
                service.Update(updatedStageCase);

            }





        }
    }
}

