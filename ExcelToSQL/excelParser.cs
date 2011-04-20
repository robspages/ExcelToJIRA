using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using JiraSoap; 
using System.Collections; 

namespace ExcelToSQL
{
    public class ExcelParser
    {
        public JSClient jira;

        public ExcelParser(JSClient jiraClient)
        {
            jira = jiraClient;    
        }


        public void BulkLoadNewIssues(ExcelSheet sheet)
        {
            if (sheet.Rows.Count != 0)
            {
                foreach (System.Data.DataRow row in sheet.Rows)
                {
                    RemoteIssue issue = this.convertRowToIssue(row, "New Feature");
                    var existingIssues = from existing in jira.Issues where existing.summary == issue.summary select existing;
                    if (new List<RemoteIssue>(existingIssues).Count == 0)
                    {
                         jira.addNewIssue(issue);
                    }
                    else
                    {
                        String existingKey = "";//existingIssues.ToArray<RemoteIssue>[0].key;
                    //    jira.updateIssue(issue);
                    }
                   
                }
            }
        }


        //slow but functional 
        public bool IssueIsDuplicate(RemoteIssue issue) 
        {
            var i = from dup in jira.Issues where dup.summary == issue.summary select dup;
            return (new List<RemoteIssue>(i).Count > 0); 
        } 

        public RemoteIssue convertRowToIssue(DataRow row, String issueType)
        {
            RemoteIssue issue = new RemoteIssue();

            issue.type = getIssueType(issueType).id;
            issue.description = BuildDescription(row); 
            issue.summary = BuildSummary(row);
            issue.components = setComponent(row["Functional Area"].ToString());
          
                RemoteVersion[] mySprint = setVersion((String.IsNullOrEmpty(row["Sprint"].ToString()) ? "Backlog" : String.Format("Sprint {0}", (row["Sprint"].ToString()))));
                issue.fixVersions = mySprint;
                issue.affectsVersions = mySprint; 
         
            return issue;
        }

        public String BuildSummary(DataRow row)
        { 
            StringBuilder sb = new StringBuilder();
            sb.AppendFormat("{0}: {1}", row["Functional Area"], row["Feature Title"]);
            if (!String.IsNullOrEmpty(row["Sub Task"].ToString()))
            {
                sb.AppendFormat(":: {0}", row["Sub Task"]);
            }
            return sb.ToString();
        }

        public String BuildDescription(DataRow row)
        {
            StringBuilder sb = new StringBuilder();
            sb.AppendLine(String.Format("Story: As a {0} I want {1} So that {2}", row["As a"], row["I Want"], row["So that"]));
            
            if (!String.IsNullOrEmpty(row["Precondition"].ToString()))
            {
                sb.AppendLine(String.Format("Precondition/Assumptions: {0}", row["Precondition"].ToString()));
            }

            sb.AppendLine(String.Format("Original Estimate: {0}, Original Dev: {1}", row["Estimate (hrs)"].ToString(), row["resource"].ToString())); 
            sb.AppendLine("Notes: " + row["Notes"].ToString());

            return sb.ToString();
        }


        /*
         * uses Linq to Objects to select the right RemoteVersion(s) from a List<RemoteVersion> taken from the Project, then convert the output to the simple array that the soap service expects
         */ 
        public RemoteVersion[] setVersion(String VersionName)
        {
            var mySprints = from sprint in jira.Versions where sprint.name == VersionName select sprint;
            return mySprints.ToArray();
        }

        /*
         * uses Linq to Objects to select the right RemoteComponent(s) from a List<RemoteComponent> taken from the Project, then convert the output to the simple array that the soap service expects
         */
        public RemoteComponent[] setComponent(String component)
        {
            var myComponents = from Component in jira.Components where Component.name == component select Component;
            return myComponents.ToArray();
        }


        public RemoteIssueType getIssueType(String myType) 
        {
            var iTypes = from types in jira.issueTypes where types.name == myType select types;
            return iTypes.ToList<RemoteIssueType>()[0]; //cast to generic list of type RemoteIssueType
        }

    }
}
