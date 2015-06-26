using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Net;
using System.IO;
using System.Xml.Schema;
using System.Xml.Xsl;
using System.Xml;
using System.Xml.XPath;
using System.Threading.Tasks;
using System.Web;
using SoaScapeLib;

namespace SoaScapeRep
{
    public class ReportUtil
    {
        #region Locals

        // IO paths

        private static string c_version = "v 2.0";
        private string m_translationsFile = "";
        string dbFileName = "";
        private string m_dataRoot = ""; // The root dierectory of the data output
        private Dictionary<string, string> m_translations = null;

        // Layer7 Monitor access
        private bool m_useFileSystem = false;
        private string m_rootURL = "";

        private string m_userid = "";
        private string m_password = "";
        private string m_domain = "";
        private bool m_windowsAuthentication = false;

        private bool hasCredentials = false; // this forces acquering new credentials
        //System.Net.CredentialCache applCredentialCache;
        System.Net.ICredentials m_credentials;

        const string NAME_SPACE_URN = "urn:soascape";
        const int MAX_QUERIES = 100;

        // XSL parameters
        private const string TOOL_VERSION = "toolVersion";
        private const string GENERATION_DATE = "generationDate";
        private const string BASE_DIR = "basedir";
        private const string TRANSFORMATION_TYPE = "type";
        private const string TRANSFORM_OUTLINE = "";
        private const string TRANSFORM_DEPLOYMENT = "deployed";
        private const string TRANSFORM_REALIZATION = "version";



        private StreamWriter m_logStream = null;

        private BusinessUnitEntity[] m_allInternalBusinessUnits;
        private BusinessUnitEntity[] m_allExternalBusinessUnits;
        private List<BusinessUnitEntity[]> m_BusinessUnits;

        private string m_lastError = "";
        #endregion

        #region Constructors

        public ReportUtil(string outputRoot, string translationsFile)
        {
            if (!Directory.Exists(outputRoot)) throw (new Exception("Data folder path: \"" + outputRoot + "\" does not exist!"));
            m_dataRoot = outputRoot;
            if (translationsFile != null)
            {
                m_translationsFile = translationsFile;
                if (!File.Exists(m_translationsFile)) throw (new Exception("Translations file: \"" + translationsFile + "\" does not exist!"));
            }
        }

        #endregion

        #region Properties

        public string LastError
        { get { return m_lastError; } }

        public string AlternativeHostURL
        {
            get { return m_rootURL; }
            set { m_rootURL = value; }
        }

        public bool UseFileSystem
        {
            get { return m_useFileSystem; }
            set { m_useFileSystem = value; }
        }

        public bool UseWindowsAuthentication
        {
            get { return m_windowsAuthentication; }
            set { m_windowsAuthentication = value; }
        }

        

        #endregion

        #region PublicFunctions

        public void ClearLastError()
        {
            m_lastError = "";
        }

        public void SetBasicAuthentication(string domain, string uid, string password)
        {
            m_domain = domain;
            m_userid = uid;
            m_password = password;
        }

        public DirectoryInfo GetLoggingDirectory()
        {
            return getOrCreateDirectory(m_dataRoot + "\\log", "Could not get Log directory: ");
        }

        public DirectoryInfo GetPublicReportsDir()
        {
            return getOrCreateDirectory(m_dataRoot + "\\reports\\public\\xml", "Could not get output directory for public reports");
        }

        public DirectoryInfo GetPrivateReportsDir(string buName)
        {
            return getOrCreateDirectory(m_dataRoot + "\\reports\\private\\" + buName + "\\xml", "Could not get output directory for public reports");
        }

        public DirectoryInfo GetSpecialReportsDir()
        {
            return getOrCreateDirectory(m_dataRoot + "\\reports\\_special\\xml", "Could not get output directory for public reports");
        }

        public DirectoryInfo GetXslDir()
        {
            return getOrCreateDirectory(m_dataRoot + "\\xsl", "Could not get directory with XSL style sheets");
        }


        public bool WriteReports(string dbPath = null, bool stopOnError = false)
        {
            bool retValue = false;
            bool endpointsLoaded = true;
            try
            {
                if (!SoaBrowser.IsLoaded())
                {
                    if (dbPath != null) retValue = SoaBrowser.LoadDB(dbPath);
                    if (!retValue) return false;
                }
                endpointsLoaded = SoaBrowser.IsRegistryLoaded;

                if (!endpointsLoaded)
                {
                    if (!m_useFileSystem)
                    {
                        if (!getCredentials()) return false;
                    }
                    foreach (ServiceDomainEntity dom in SoaBrowser.AllServiceDomains)
                        foreach (EndpointEntity ep in dom.AllMediatedEndpoints) ep.SetProtectedEndpoint(getProtectedEndpoint(ep.UrlRegistryFolder));
                }
                
                DirectoryInfo outDir;
                DirectoryInfo logDir = GetLoggingDirectory();
                m_logStream = createLogFile(logDir, "reportLog");
                if ((m_translationsFile.Length > 0) && File.Exists(m_translationsFile)) fillTranslationsDictionary();
                m_allInternalBusinessUnits = SoaBrowser.AllInternalBusinessUnits;
                m_allExternalBusinessUnits = SoaBrowser.AllExternalBusinessUnits;
                m_BusinessUnits = new List<BusinessUnitEntity[]> { m_allInternalBusinessUnits, m_allExternalBusinessUnits };

                // WRITE PUBLIC
                outDir = GetPublicReportsDir();
                if (outDir == null) return false;
                repApplications(outDir);
                repApplicationVersions(outDir);
                repDeployedApplications(outDir);
                repServices(outDir);
                repServiceVersions(outDir);
                repDeployedServices(outDir);

                // WRITE SPECIAL
                outDir = GetSpecialReportsDir();
                if (outDir == null) return false;
                repMediatedEndpoints(outDir);
                repServicesAndConsumers(outDir);
                repConsumersAndServices(outDir);

                // WRITE PRIVATE
                foreach (BusinessUnitEntity bu in m_allInternalBusinessUnits)
                {
                    outDir = GetPrivateReportsDir(bu.Name);
                    if (outDir != null)
                    {
                        repMediatedEndpoints(outDir, false, bu);
                        repServicesAndConsumers(outDir, false, bu);
                        repConsumersAndServices(outDir, false, bu);
                    }
                }
            }
            finally
            {
                if (m_logStream != null) m_logStream.Close();
                m_logStream = null;
            }
            return true;
        }

        #endregion

        #region PrivateFunctions

        private bool repServices(DirectoryInfo outDir)
        {
            string xlsFilename = "services.xsl";
            string repName = "Services";
            try
            {
                FileStream file = createReportFile(outDir, repName, "");
                XmlWriter wr = writeReportPreamble(repName, file);
                bool reportOK = writeServices(wr);
                writePostambleAndClose(file, wr);
                if ((reportOK) && (xlsFilename != null)) transformToHTML(outDir, repName, file.Name, xlsFilename, false);
            }
            catch (Exception ex)
            {
                string m_lastError = "Error occured: " + ex.Message;
                if (ex.InnerException != null) m_lastError += ex.InnerException.Message;
                if (m_logStream != null) m_logStream.WriteLine(m_lastError);
            }
            return true;
        }

        private bool repServiceVersions(DirectoryInfo outDir)
        {
            string xlsFilename = "services.xsl";
            string repName = "Services-Versions";
            try
            {
                FileStream file = createReportFile(outDir, repName, "");
                XmlWriter wr = writeReportPreamble(repName, file);
                bool reportOK = writeServiceVersions(wr);
                writePostambleAndClose(file, wr);
                if ((reportOK) && (xlsFilename != null)) transformToHTML(outDir, repName, file.Name, xlsFilename, false, TRANSFORM_REALIZATION);
            }
            catch (Exception ex)
            {
                string m_lastError = "Error occured: " + ex.Message;
                if (ex.InnerException != null) m_lastError += ex.InnerException.Message;
                if (m_logStream != null) m_logStream.WriteLine(m_lastError);
            }
            return true;
        }

        private bool repDeployedServices(DirectoryInfo outDir)
        {
            string xlsFilename = "services.xsl";
            string repName = "Services-Deployed";
            try
            {
                FileStream file = createReportFile(outDir, repName, "");
                XmlWriter wr = writeReportPreamble(repName, file);
                bool reportOK = writeDeployedServices(wr);
                writePostambleAndClose(file, wr);
                if ((reportOK) && (xlsFilename != null)) transformToHTML(outDir, repName, file.Name, xlsFilename, false, TRANSFORM_DEPLOYMENT);
            }
            catch (Exception ex)
            {
                string m_lastError = "Error occured: " + ex.Message;
                if (ex.InnerException != null) m_lastError += ex.InnerException.Message;
                if (m_logStream != null) m_logStream.WriteLine(m_lastError);
            }
            return true;
        }


        private bool repDeployedApplications(DirectoryInfo outDir)
        {
            string xlsFilename = "applications.xsl";
            string repName = "Applications-Deployed";
            try
            {
                FileStream file = createReportFile(outDir, repName, "");
                XmlWriter wr = writeReportPreamble(repName, file);
                bool reportOK = writeDeployedApplications(wr);
                writePostambleAndClose(file, wr);
                if ((reportOK) && (xlsFilename != null)) transformToHTML(outDir, repName, file.Name, xlsFilename, false, TRANSFORM_DEPLOYMENT);
            }
            catch (Exception ex)
            {
                string m_lastError = "Error occured: " + ex.Message;
                if (ex.InnerException != null) m_lastError += ex.InnerException.Message;
                if (m_logStream != null) m_logStream.WriteLine(m_lastError);
            }
            return true;
        }

        private bool repApplicationVersions(DirectoryInfo outDir)
        {
            string xlsFilename = "applications.xsl";
            string repName = "Applications-Versions";
            try
            {
                FileStream file = createReportFile(outDir, repName, "");
                XmlWriter wr = writeReportPreamble(repName, file);
                bool reportOK = writeApplicationVersions(wr);
                writePostambleAndClose(file, wr);
                if ((reportOK) && (xlsFilename != null)) transformToHTML(outDir, repName, file.Name, xlsFilename, false, TRANSFORM_REALIZATION);
            }
            catch (Exception ex)
            {
                string m_lastError = "Error occured: " + ex.Message;
                if (ex.InnerException != null) m_lastError += ex.InnerException.Message;
                if (m_logStream != null) m_logStream.WriteLine(m_lastError);
            }
            return true;
        }

        private bool repApplications(DirectoryInfo outDir)
        {
            string xlsFilename = "applications.xsl";
            string repName = "Applications";
            try
            {
                FileStream file = createReportFile(outDir, repName, "");
                XmlWriter wr = writeReportPreamble(repName, file);
                bool reportOK = writeApplications(wr);
                writePostambleAndClose(file, wr);
                if ((reportOK) && (xlsFilename != null)) transformToHTML(outDir, repName, file.Name, xlsFilename, false);
            }
            catch (Exception ex)
            {
                string m_lastError = "Error occured: " + ex.Message;
                if (ex.InnerException != null) m_lastError += ex.InnerException.Message;
                if (m_logStream != null) m_logStream.WriteLine(m_lastError);
            }
            return true;
        }


        private bool repMediatedEndpoints(DirectoryInfo outDir, bool daily = false, BusinessUnitEntity bu = null)
        {
            string xlsFilename = "endpoints.xsl";
            string repName = "Endpoints-Mediated";
            string fileAddon = daily ? "_" + DateTime.Now.ToShortDateString() : "";
            bool isPrivate = bu != null;
            try
            {
                FileStream file = createReportFile(outDir, repName, fileAddon);
                XmlWriter wr = writeReportPreamble(repName, file);
                bool reportOK = writeEndpoints(wr, bu);
                writePostambleAndClose(file, wr);
                if ((reportOK) && (xlsFilename != null)) transformToHTML(outDir, repName, file.Name, xlsFilename, isPrivate);
            }
            catch (Exception ex)
            {
                string m_lastError = "Error occured: " + ex.Message;
                if (ex.InnerException != null) m_lastError += ex.InnerException.Message;
                if (m_logStream != null) m_logStream.WriteLine(m_lastError);
            }
            return true;
        }

        private bool repServicesAndConsumers(DirectoryInfo outDir, bool daily = false, BusinessUnitEntity bu = null)
        {
            string xlsFilename = "consumed-per-service.xsl";
            string fileAddon = daily ? "_" + DateTime.Now.ToShortDateString() : "";
            bool isPrivate = bu != null;
            try
            {
                string repName = "Services-ConsumedPerService";
                FileStream file = createReportFile(outDir, repName, fileAddon);
                XmlWriter wr = writeReportPreamble(repName, file);
                bool reportOK = writeServicesAndConsumers(wr, bu);
                writePostambleAndClose(file, wr);
                if ((reportOK) && (xlsFilename != null)) transformToHTML(outDir, repName, file.Name, xlsFilename, isPrivate);
            }
            catch (Exception ex)
            {
                string m_lastError = "Error occured: " + ex.Message;
                if (ex.InnerException != null) m_lastError += ex.InnerException.Message;
                if (m_logStream != null) m_logStream.WriteLine(m_lastError);
            }
            return true;
        }


        private bool repConsumersAndServices(DirectoryInfo outDir, bool daily = false, BusinessUnitEntity bu = null)
        {
            string xlsFilename = "consumed-per-application.xsl";
            string fileAddon = daily ? "_" + DateTime.Now.ToShortDateString() : "";
            bool isPrivate = bu != null;
            try
            {
                string repName = "Services-ConsumedPerApplication";
                FileStream file = createReportFile(outDir, repName, fileAddon);
                XmlWriter wr = writeReportPreamble(repName, file);
                bool reportOK = writeConsumersAndServices(wr, bu);
                writePostambleAndClose(file, wr);
                if ((reportOK) && (xlsFilename != null)) transformToHTML(outDir, repName, file.Name, xlsFilename, isPrivate);
            }
            catch (Exception ex)
            {
                string m_lastError = "Error occured: " + ex.Message;
                if (ex.InnerException != null) m_lastError += ex.InnerException.Message;
                if (m_logStream != null) m_logStream.WriteLine(m_lastError);
            }
            return true;
        }

        private bool writeApplications(XmlWriter wr)
        {
            for (int i = 0; i <= 1; i++)
            {
                BusinessUnitEntity[] buArray = m_BusinessUnits[i];
                string buNodeName = (i == 0) ? "BusinessUnit" : "ExternalOrganization";
                foreach (BusinessUnitEntity bu in m_BusinessUnits[i])
                {
                    wr.WriteStartElement(buNodeName);
                    string buName = (i == 0) ? bu.Name : bu.Organization.Name;
                    wr.WriteAttributeString("name", buName);
                    wr.WriteStartElement("Applications"); 
                    foreach (ApplEntity appl in bu.Applications)
                    {
                        wr.WriteStartElement("Application");
                        wr.WriteElementString("name", appl.Name);
                        wr.WriteElementString("description", appl.Description);
                        wr.WriteEndElement();
                    }
                    wr.WriteEndElement();
                    wr.WriteEndElement();
                }
            }
            return true;
        }

        private bool writeApplicationVersions(XmlWriter wr)
        {
            for (int i = 0; i <= 1; i++)
            {
                BusinessUnitEntity[] buArray = m_BusinessUnits[i];
                string buNodeName = (i == 0) ? "BusinessUnit" : "ExternalOrganization";
                foreach (BusinessUnitEntity bu in m_BusinessUnits[i])
                {
                    wr.WriteStartElement(buNodeName);
                    string buName = (i == 0) ? bu.Name : bu.Organization.Name;
                    wr.WriteAttributeString("name", buName);
                    wr.WriteStartElement("Applications");
                    foreach (ApplEntity appl in bu.Applications)
                    {
                        foreach (ApplEntity applVersion in appl.Originated)
                        {
                            wr.WriteStartElement("Application");
                            wr.WriteElementString("name", appl.Name);
                            wr.WriteElementString("version", applVersion.Version);
                            wr.WriteElementString("description", appl.Description);
                            wr.WriteElementString("versionDescription", applVersion.Description);
                            wr.WriteEndElement();
                        }
                    }
                    wr.WriteEndElement();
                    wr.WriteEndElement();
                }
            }
            return true;
        }


        private bool writeDeployedApplications(XmlWriter wr)
        {
            foreach (ServiceDomainEntity dom in SoaBrowser.AllServiceDomains)
            {
                wr.WriteStartElement("ServiceDomain");
                wr.WriteAttributeString("name", dom.Name);
                for (int i = 0; i <= 1; i++)
                {
                    BusinessUnitEntity[] buArray = m_BusinessUnits[i];
                    string buNodeName = (i == 0) ? "BusinessUnit" : "ExternalOrganization";
                    foreach (BusinessUnitEntity bu in m_BusinessUnits[i])
                    {
                        wr.WriteStartElement(buNodeName); // **** "BusinessUnits" / "ExternalOrganization"
                        string buName = (i == 0) ? bu.Name : bu.Organization.Name;
                        wr.WriteAttributeString("name", buName);
                        wr.WriteStartElement("Applications"); // **** "Applications"
                        foreach (ApplEntity appl in bu.DeployedAppsInDomain(dom))
                        {
                            wr.WriteStartElement("Application");
                            wr.WriteElementString("name", appl.DisplayNameReport);
                            wr.WriteElementString("version", appl.Version);
                            wr.WriteElementString("description", appl.Origin.Origin.Description);
                            wr.WriteElementString("versionDescription", appl.Origin.Description);
                            wr.WriteElementString("deploymentDescription", appl.Description);
                            wr.WriteEndElement(); // Application
                        }
                        wr.WriteEndElement(); // Applications
                        wr.WriteEndElement(); // "BusinessUnits" / "ExternalOrganization"
                    }
                }
                wr.WriteEndElement(); // ServiceDomain
            }
            return true;
        }

        private bool writeServices(XmlWriter wr)
        {
            for (int i = 0; i <= 1; i++)
            {
                BusinessUnitEntity[] buArray = m_BusinessUnits[i];
                string buNodeName = (i == 0) ? "BusinessUnit" : "ExternalOrganization";
                foreach (BusinessUnitEntity bu in m_BusinessUnits[i])
                {
                    wr.WriteStartElement(buNodeName);
                    string buName = (i == 0) ? bu.Name : bu.Organization.Name;
                    wr.WriteAttributeString("name", buName);
                    wr.WriteStartElement("Services");
                    foreach (ApplEntity appl in bu.Applications)
                    {
                        foreach (ServiceEntity svc in appl.Services)
                        {
                            wr.WriteStartElement("Service");
                            wr.WriteElementString("name", svc.Name);
                            wr.WriteElementString("description", svc.Description);
                            wr.WriteElementString("providingApplication", appl.Name);
                            wr.WriteEndElement();
                        }
                    }
                    wr.WriteEndElement();
                    wr.WriteEndElement();
                }
            }
            return true;
        }

        private bool writeServiceVersions(XmlWriter wr)
        {
            for (int i = 0; i <= 1; i++)
            {
                BusinessUnitEntity[] buArray = m_BusinessUnits[i];
                string buNodeName = (i == 0) ? "BusinessUnit" : "ExternalOrganization";
                foreach (BusinessUnitEntity bu in m_BusinessUnits[i])
                {
                    wr.WriteStartElement(buNodeName);
                    string buName = (i == 0) ? bu.Name : bu.Organization.Name;
                    wr.WriteAttributeString("name", buName);
                    wr.WriteStartElement("Services");
                    foreach (ApplEntity appl in bu.Applications)
                    {
                        foreach (ApplEntity applVersion in appl.Originated)
                        {
                            foreach (ServiceEntity svc in applVersion.Services)
                            {
                                wr.WriteStartElement("Service");
                                wr.WriteElementString("name", svc.Name);
                                wr.WriteElementString("version", svc.Version);
                                wr.WriteElementString("description", svc.Description);
                                wr.WriteElementString("providingApplication", applVersion.Name);
                                wr.WriteElementString("applicationVersion", applVersion.Version);
                                wr.WriteEndElement();
                            }
                        }
                    }
                    wr.WriteEndElement();
                    wr.WriteEndElement();
                }
            }
            return true;
        }


        private bool writeDeployedServices(XmlWriter wr)
        {
            foreach (ServiceDomainEntity dom in SoaBrowser.AllServiceDomains)
            {
                wr.WriteStartElement("ServiceDomain");
                wr.WriteAttributeString("name", dom.Name);
                for (int i = 0; i <= 1; i++)
                {
                    BusinessUnitEntity[] buArray = m_BusinessUnits[i];
                    string buNodeName = (i == 0) ? "BusinessUnit" : "ExternalOrganization";
                    foreach (BusinessUnitEntity bu in m_BusinessUnits[i])
                    {
                        wr.WriteStartElement(buNodeName);
                        string buName = (i == 0) ? bu.Name : bu.Organization.Name;
                        wr.WriteAttributeString("name", buName);
                        wr.WriteStartElement("Services"); // "Services"
                        foreach (ApplEntity appl in bu.DeployedAppsInDomain(dom))
                        {
                            foreach (ServiceEntity svc in appl.DeployedServices)
                            {
                                wr.WriteStartElement("Service");
                                wr.WriteElementString("providingApplication", appl.DisplayNameReport);
                                wr.WriteElementString("applicationVersion", appl.Version);
                                wr.WriteElementString("name", svc.Name);
                                wr.WriteElementString("version", svc.Version);
                                wr.WriteElementString("description", svc.Origin.Origin.Description);
                                wr.WriteElementString("versionDescription", svc.Origin.Description);
                                wr.WriteElementString("deploymentDescription", svc.Description);
                                wr.WriteElementString("targetNamespace", svc.TargetNamespace);
                                wr.WriteElementString("definitionName", svc.DefinitionName);
                                wr.WriteEndElement(); // Service
                            }
                        }
                        wr.WriteEndElement(); // Applications
                        wr.WriteEndElement(); // BusinessUnit
                    }
                }
                wr.WriteEndElement(); // ServiceDomain
            }
            return true;
        }


        private bool writeEndpoints(XmlWriter wr, BusinessUnitEntity businessUnit = null)
        {
            foreach (ServiceDomainEntity dom in SoaBrowser.AllServiceDomains)
            {
                wr.WriteStartElement("ServiceDomain");
                wr.WriteAttributeString("name", dom.Name);
                for (int i = 0; i <= 1; i++)
                {
                    BusinessUnitEntity[] buArray = m_BusinessUnits[i];
                    string buNodeName = (i == 0) ? "BusinessUnit" : "ExternalOrganization";
                    foreach (BusinessUnitEntity bu in m_BusinessUnits[i])
                    {
                        if ((businessUnit != null) && (businessUnit != bu)) continue;  // do only one if specified
                        wr.WriteStartElement(buNodeName);
                        string buName = (i == 0) ? bu.Name : bu.Organization.Name;
                        wr.WriteAttributeString("name", buName);
                        wr.WriteStartElement("Services"); // "Services"
                        foreach(ApplEntity appl in bu.DeployedAppsInDomain(dom))
                        {
                            foreach (ServiceEntity srv in appl.DeployedServices)
                            {
                                wr.WriteStartElement("Service");
                                wr.WriteAttributeString("name", srv.Name);
                                wr.WriteAttributeString("application", appl.DisplayNameReport);
                                wr.WriteStartElement("Endpoints");
                                foreach (EndpointEntity ep in srv.Endpoints)  ep.SerializeInXml(wr, "Endpoint", appl.Version, srv.Version);
                                wr.WriteEndElement();
                                wr.WriteEndElement();
                            }
                        }
                        wr.WriteEndElement();
                        wr.WriteEndElement();
                    }
                }
                wr.WriteEndElement();
            }
            return true;
        }


        private bool writeServicesAndConsumers(XmlWriter wr, BusinessUnitEntity businessUnit = null)
        {
            foreach (ServiceDomainEntity dom in SoaBrowser.AllServiceDomains)
            {
                wr.WriteStartElement("ServiceDomain");
                wr.WriteAttributeString("name", dom.Name);
                for (int i = 0; i <= 1; i++)
                {
                    BusinessUnitEntity[] buArray = m_BusinessUnits[i];
                    string buNodeName = (i == 0) ? "BusinessUnit" : "ExternalOrganization";
                    foreach (BusinessUnitEntity bu in m_BusinessUnits[i])
                    {
                        if ((businessUnit != null) && (businessUnit != bu)) continue;  // do only one if specified
                        wr.WriteStartElement(buNodeName);
                        string buName = (i == 0) ? bu.Name : bu.Organization.Name;
                        wr.WriteAttributeString("name", buName);
                        wr.WriteStartElement("Services"); // "Services"
                        foreach (ApplEntity appl in bu.DeployedAppsInDomain(dom))
                        {
                            foreach (ServiceEntity deployedSrv in appl.DeployedServices)
                            {
                                wr.WriteStartElement("Service");
                                wr.WriteAttributeString("name", deployedSrv.Name);
                                wr.WriteAttributeString("application", appl.DisplayNameReport);
                                wr.WriteStartElement("Applications");
                                foreach (EndpointEntity ep in deployedSrv.Endpoints)
                                {
                                    foreach (ApplEntity consumingAppl in ep.ConsumedByApplications)
                                    {
                                        wr.WriteStartElement("Application");
                                        wr.WriteElementString("name", consumingAppl.DisplayNameReport);
                                        wr.WriteElementString("version", consumingAppl.Version);
                                        wr.WriteElementString("endpointName", ep.IntermediaryGivenName);
                                        wr.WriteElementString("gateway", ep.Intermediary.Name);
                                        wr.WriteElementString("resolutionPath", ep.ResolutionPath);
                                        wr.WriteElementString("targetNamespace", deployedSrv.TargetNamespace);
                                        wr.WriteEndElement();
                                    }
                                }
                                wr.WriteEndElement();
                                wr.WriteEndElement();
                            }
                        }
                        wr.WriteEndElement();
                        wr.WriteEndElement();
                    }
                }
                wr.WriteEndElement();
            }
            return true;
        }

        private bool writeConsumersAndServices(XmlWriter wr, BusinessUnitEntity businessUnit = null)
        {
            foreach (ServiceDomainEntity dom in SoaBrowser.AllServiceDomains)
            {
                wr.WriteStartElement("ServiceDomain");
                wr.WriteAttributeString("name", dom.Name);
                for (int i = 0; i <= 1; i++)
                {
                    BusinessUnitEntity[] buArray = m_BusinessUnits[i];
                    string buNodeName = (i == 0) ? "BusinessUnit" : "ExternalOrganization";
                    foreach (BusinessUnitEntity bu in m_BusinessUnits[i])
                    {
                        if ((businessUnit != null) && (businessUnit != bu)) continue;  // do only one if specified
                        wr.WriteStartElement(buNodeName);
                        string buName = (i == 0) ? bu.Name : bu.Organization.Name;
                        wr.WriteAttributeString("name", buName);
                        wr.WriteStartElement("Applications"); // "Services"
                        foreach (ApplEntity appl in bu.DeployedAppsInDomain(dom))
                        {
                            wr.WriteStartElement("Application");
                            wr.WriteAttributeString("name", appl.DisplayNameReport);
                            wr.WriteStartElement("Services");
                            foreach (EndpointEntity ep in appl.ConsumedEndpoints)
                            {
                                wr.WriteStartElement("Service");
                                wr.WriteElementString("consumingApplicationVersion", appl.Version);
                                wr.WriteElementString("name", ep.ServiceName);
                                wr.WriteElementString("version", ep.Parent.Version);
                                wr.WriteElementString("targetNamespace", ep.Parent.TargetNamespace);
                                wr.WriteElementString("endpointName", ep.IntermediaryGivenName);
                                wr.WriteElementString("gateway", ep.Intermediary.Name);
                                wr.WriteElementString("resolutionPath", ep.ResolutionPath);
                                wr.WriteElementString("application", ep.ProvidingApplication.DisplayNameReport);
                                wr.WriteElementString("applicationVersion", ep.ProvidingApplication.Version);
                                wr.WriteElementString("businessUnit", ep.ProvidingApplication.BusinessUnit.Name);
                                wr.WriteElementString("organization", ep.ProvidingApplication.BusinessUnit.Organization.Name);
                                wr.WriteEndElement(); // Service
                            }
                            wr.WriteEndElement(); // Services
                            wr.WriteEndElement(); // Application
                        }
                        wr.WriteEndElement(); // Applications
                        wr.WriteEndElement(); // BusinessUnit
                    }
                }
                wr.WriteEndElement(); // ServiceDomain
            }
            return true;
        }


        private FileStream createReportFile(DirectoryInfo outDir, string repName, string fileAddon)
        {
            string folderPath = outDir.FullName + "\\";
            string fileName = folderPath + repName + fileAddon + ".xml";
            FileStream file = new FileStream(fileName, FileMode.Create);
            return file;
        }

        private XmlWriter writeReportPreamble(string repName, FileStream file)
        {
            XmlWriterSettings writerSettings = new XmlWriterSettings();
            writerSettings.Indent = true;
            writerSettings.NewLineOnAttributes = false;
            writerSettings.Encoding = new UTF8Encoding();
            writerSettings.OmitXmlDeclaration = true;
            XmlWriter wr = XmlWriter.Create(file, writerSettings);
            wr.WriteStartDocument();
            wr.WriteStartElement(reportName2StartElement(repName), NAME_SPACE_URN);
            wr.WriteAttributeString("xmlns", null, null, NAME_SPACE_URN);
            return wr;
        }

        private void writePostambleAndClose(FileStream file, XmlWriter wr)
        {
            wr.WriteEndElement();
            wr.Flush();
            wr.Close();
            file.Close();
        }


        private DirectoryInfo getOrCreateDirectory(string path, string errorText)
        {
            try
            {
                DirectoryInfo outputDir = new DirectoryInfo(path);
                if (!outputDir.Exists) outputDir.Create();
                return outputDir;
            }
            catch (Exception ex)
            {
                m_lastError = errorText + path + " " + ex.Message.ToString();
                if (m_logStream != null) m_logStream.WriteLine(m_lastError);
                return null;
            }
        }

        void fillTranslationsDictionary()
        {
            if (!File.Exists(m_translationsFile)) return;
            XmlDocument trans = new XmlDocument();
            try
            {
                trans.Load(m_translationsFile);
                m_translations = new Dictionary<string, string>();
                XmlNamespaceManager nsmgr = new XmlNamespaceManager(trans.NameTable);
                //nsmgr.AddNamespace("", "urn:soascape");
                foreach (XmlNode translationPair in trans.DocumentElement.ChildNodes)
                {
                    XmlNode fromNode = translationPair.SelectSingleNode("From");
                    XmlNode toNode = translationPair.SelectSingleNode("To");
                    if ((fromNode != null) && (toNode != null)) m_translations.Add(fromNode.InnerText, toNode.InnerText);
                }
            }
            catch (Exception ex)
            {
                string m_lastError = "Error occured: " + ex.Message;
                if (ex.InnerException != null) m_lastError += ex.InnerException.Message;
                if (m_logStream != null) m_logStream.WriteLine(m_lastError);
                m_translations = null;
            }

        }


        private DataSet executeQuery(string command, string datasetName = "DS", string tableName = "TABLE")
        {
            if ((dbFileName == null) || (dbFileName.Length == 0))
            {
                m_lastError = "Internal error: database file missing (function executeQuery)!";
                return null;
            }
            string connectionString = "Data Source=" + dbFileName + ";Version=3;";
            UtilDB dbUtil = new UtilDB(connectionString);
            if (!(dbUtil.Load(command, "")))
            {
                m_lastError = dbUtil.LastError;
                return null;
            }
            DataSet ds = dbUtil.DataSet;
            ds.DataSetName = datasetName;
            ds.Namespace = NAME_SPACE_URN;
            ds.Tables[0].TableName = tableName;
            return ds;
        }

        private XmlDocument LoadQueries(string fileName)
        {
            if ((fileName == null) || (fileName.Length == 0))
            {
                m_lastError = "Internal error: query filename zero length";
                return null;
            }
            try
            {
                XmlDocument queriesXML = new XmlDocument();
                queriesXML.Load(fileName);
                return queriesXML;
            }
            catch (Exception ex)
            {
                m_lastError = "Load failed, file: \"" + fileName + "\" >>" + ex.Message;
                if (ex.InnerException != null) m_lastError += ex.InnerException.Message;
                if (m_logStream != null) m_logStream.WriteLine(m_lastError);
                return null;
            }
        }


        // xmlDir is the directory with xml file
        // fileNameXML is the full filename of the input xmlFile
        // queryName is the name of the query
        // xslFile is the filename only, including the extension of the xsl stylesheer
        private void transformToHTML(DirectoryInfo xmlDir, string queryName, string fileNameXML, string xslFile, bool isPrivate, string transformtionType = "")
        {
            string fileNameXsl = m_dataRoot + "\\xsl\\" + xslFile;
            if (!File.Exists(fileNameXsl)) return;
            StreamWriter htmlOutput = null;
            try
            {
                DirectoryInfo outputDir = xmlDir.Parent;
                string fileNameHTML = outputDir.FullName + "\\" + query2NameHTML(queryName);
                var myXslTrans = new XslCompiledTransform();
                XsltArgumentList argList = new XsltArgumentList();
                argList.AddParam("toolVersion", "", c_version);
                argList.AddParam("generationDate", "", DateTime.Now.ToLongDateString());
                if ((transformtionType.Equals(TRANSFORM_DEPLOYMENT)) || (transformtionType.Equals(TRANSFORM_REALIZATION)))
                    argList.AddParam("type", "", transformtionType);
                if (isPrivate) argList.AddParam("basedir", "", "./../../../"); else argList.AddParam("basedir", "", "./../../");
                myXslTrans.Load(fileNameXsl);
                htmlOutput = File.CreateText(fileNameHTML);
                myXslTrans.Transform(fileNameXML, argList, htmlOutput);
            }
            catch (Exception ex)
            {
                string m_lastError = "Error occured: " + ex.Message;
                if (ex.InnerException != null) m_lastError += ex.InnerException.Message;
                if (m_logStream != null) m_logStream.WriteLine(m_lastError);
            }
            finally
            {
                if ((htmlOutput != null) && (htmlOutput.BaseStream != null)) htmlOutput.Close();
            }
        }







        string reportName2StartElement(string queryName)
        {
            return reportName2Name(queryName) + "s";
        }

        // Return the name of the HTML output file 
        string query2NameHTML(string queryName)
        {
            int idx1, idx2;
            string translatedQueryName = queryName;
            foreach (string key in m_translations.Keys)
            {
                translatedQueryName = translatedQueryName.Replace(key, m_translations[key]);
            }
            idx1 = translatedQueryName.IndexOf('_');
            idx2 = translatedQueryName.IndexOf('-');
            if (idx1 < 0) return translatedQueryName + ".html";
            string part1 = translatedQueryName.Substring(0, idx1);
            if (idx2 < 0) return part1 + ".html";
            int lastCharIdx = translatedQueryName.Length - 1;
            if (idx2 == lastCharIdx) return part1;
            return part1 + translatedQueryName.Substring(idx2) + ".html";
        }


        // Create a suitable name for the XML element that is generated from the recordset name
        string reportName2Name(string repName)
        {
            int idx;
            idx = repName.IndexOf('-');
            if (idx > 1) return repName.Substring(0, idx - 1);
            idx = repName.Length - 1;
            if (repName[idx] == 's')
            {
                return repName.Substring(0, idx);
            }
            else
            {
                int idx2;
                for (idx2 = 1; idx2 <= idx; idx2++)
                {
                    if (Char.IsUpper(repName[idx2])) break;
                }
                return repName.Substring(0, idx2);
            }
        }



        private string replaceHost(string urlPath)
        {
            string result = urlPath;
            try
            {
                int count = m_rootURL.Split('/').Length - 1;
                string addOn = urlPath.Substring(indexOfNth(urlPath, '/', count) + 1);
                result = m_rootURL + addOn;
            }
            catch (Exception ex)
            { // some logging
                string err = "Failed to replace host in: " + urlPath + " (" + m_rootURL + " \n " + ex.Message.ToString();
                if (ex.InnerException != null) err += ex.InnerException.Message;
                if (m_logStream != null) m_logStream.WriteLine(err);
            }
            return result;
        }


        private string getProtectedEndpoint(string urlRegistryFolder)
        {
            string retValue = "";
            if ((urlRegistryFolder == null) || (urlRegistryFolder.Length < 3)) return retValue;
            try
            {
                string urlRegistryFolder2 = urlRegistryFolder;
                if (m_rootURL.Length > 0) urlRegistryFolder2 = replaceHost(urlRegistryFolder);
                XmlDocument doc = new XmlDocument();
                if (m_useFileSystem)
                {
                    string physicalUrlFilePath;
                    if (System.Web.HttpContext.Current == null) // This means the library is invoked from a desktop application
                    {
                        physicalUrlFilePath = urlRegistryFolder2.Replace('/', '\\') + "index.xml";
                    }
                    else
                    {
                        urlRegistryFolder2 = removeHost(urlRegistryFolder2);
                        // http://stackoverflow.com/questions/1190196/using-server-mappath-in-external-c-sharp-classes-in-asp-net
                        physicalUrlFilePath = System.Web.HttpContext.Current.Server.MapPath(urlRegistryFolder2 + "index.xml");
                    }
                    string physicalFilePath = HttpUtility.UrlDecode(physicalUrlFilePath);
                    doc.Load(physicalFilePath);
                }
                else
                {
                    XmlUrlResolver resolver = new XmlUrlResolver();
                    ServicePointManager.ServerCertificateValidationCallback = delegate { return true; };
                    resolver.Credentials = m_credentials;
                    doc.XmlResolver = resolver;
                    doc.Load(urlRegistryFolder2 + "index.xml");
                }
                XmlNamespaceManager ns = new XmlNamespaceManager(doc.NameTable); // Fuck SelectSingleNode!
                XmlElement root = doc.DocumentElement;
                foreach (XmlNode aNode in root.ChildNodes)
                {
                    if (aNode.Name.Equals("protectedEndpoint"))
                    {
                        retValue = aNode.InnerText;
                        break;
                    }
                }
            }
            catch (Exception ex)
            { // some logging
                string err = "Failed reading index.xml from: " + urlRegistryFolder + "\n " + ex.Message.ToString();
                if (ex.InnerException != null) err += ex.InnerException.Message;
                if (m_logStream != null) m_logStream.WriteLine(err);
            }
            return retValue;
        }


        private bool getCredentials()
        {
            try
            {
                if (m_windowsAuthentication)
                    m_credentials = System.Net.CredentialCache.DefaultCredentials;
                else
                    m_credentials = new System.Net.NetworkCredential(m_userid, m_password, m_domain);
                hasCredentials = true;
            }
            catch (Exception ex)
            {
                string err;
                if (m_windowsAuthentication)
                    err = "ERR: Unable to get credentials for Windows Authentication \n" + ex.Message;
                else
                    err = "ERR: Unable to get credentials for userid: " + m_userid + " domain: " + m_domain + "\n" + ex.Message;

                if (ex.InnerException != null) err += ex.InnerException.Message;
                if (m_logStream != null) m_logStream.WriteLine(err);
                hasCredentials = false;
            }
            return hasCredentials;
        }


        //  
        private string removeHost(string urlPath)
        {
            string result = urlPath;
            try
            {
                int idx = urlPath.IndexOf(':');
                if ((idx < 0) || ((urlPath.Length - idx) < 4)) return urlPath; // nothing to remove
                int idxStart = urlPath.IndexOf('/', idx + 3);
                result = urlPath.Substring(idxStart);
            }
            catch (Exception ex)
            { // some logging
                string err = "Failed to remove the host part from: " + urlPath + ex.Message.ToString();
                if (ex.InnerException != null) err += ex.InnerException.Message;
                if (m_logStream != null) m_logStream.WriteLine(err);
            }
            return result;
        }


        private int indexOfNth(string str, char c, int n)
        {
            int remaining = n;
            for (int i = 0; i < str.Length; i++)
            {
                if (str[i] == c)
                {
                    remaining--;
                    if (remaining == 0)
                    {
                        return i;
                    }
                }
            }
            return -1;
        }

        private StreamWriter createLogFile(DirectoryInfo dirInfo, string logfilePrefix)
        {
            if (dirInfo == null) return null;
            StreamWriter outputStream;
            try
            {
                string timeStamp = ((DateTime.Now.ToFileTime() - (new DateTime(2014, 4, 8, 0, 0, 0)).ToFileTime()) / 100000).ToString();
                string destPath = System.IO.Path.Combine(dirInfo.FullName, logfilePrefix + timeStamp + ".txt");
                FileInfo f = new FileInfo(destPath);
                outputStream = f.CreateText();
                outputStream.WriteLine("Log open " + DateTime.Now.ToLongTimeString());
            }
            catch (Exception ex)
            {
                m_lastError = "Could not create the log file in: " + dirInfo.FullName + ex.Message.ToString();
                outputStream = null;
            }
            return outputStream;
        }


        #endregion

    }
}
