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
        
        public ExcelParser()
        {
           
        }

        //Sprint	resource	Functional Area Priority (higher # = lower priority)	Function priority (1 = high, 2 = medium, 3 = low)	Functional Area	Feature Title	Sub Task	As a	I want	So that	Precondition	Estimate (hrs)	Notes							

        public RemoteIssue convertRowToIssue(DataRow row, String issueType, JSClient jira)
        {
            RemoteIssue issue = new RemoteIssue();

            issue.type = getIssueType(issueType, jira).id;
            issue.description = BuildDescription(row); 
            issue.summary = BuildSummary(row);
            issue.components = setComponent(row["Functional Area"].ToString(), jira);
            if (!String.IsNullOrEmpty(row["Sprint"].ToString()))
            {
                issue.affectsVersions = setSprint(int.Parse(row["Sprint"].ToString()), jira);
            }

            return jira.addNewIssue(issue);
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
            sb.AppendLine(String.Format("As a {0} I want {1} So that {2}", row["As a"], row["I Want"], row["So that"]));
            
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
        public RemoteVersion[] setSprint(int sprintNumber, JSClient jira)
        {
           var mySprints = from sprint in jira.Versions
                         where sprint.name == String.Format("Sprint {0}", sprintNumber)
                         select sprint;

            return mySprints.ToArray();
        }


        /*
         * uses Linq to Objects to select the right RemoteComponent(s) from a List<RemoteComponent> taken from the Project, then convert the output to the simple array that the soap service expects
         */
        public RemoteComponent[] setComponent(String component, JSClient jira)
        {
            var myComponents = from Component in jira.Components
                                                 where Component.name == component
                                                 select Component;

            return myComponents.ToArray();
        }


        public RemoteIssueType getIssueType(String myType, JSClient jira) 
        {
            var iTypes = from types in jira.issueTypes
                         where types.name == myType
                         select types;

            return iTypes.ToList<RemoteIssueType>()[0]; //cast to generic list of type RemoteIssueType
        }

    }
}
