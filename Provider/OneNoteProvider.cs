#region Microsoft Community License
/*****
Microsoft Community License (Ms-CL)
Published: October 12, 2006

   This license governs  use of the  accompanying software. If you use
   the  software, you accept this  license. If you  do  not accept the
   license, do not use the software.

1. Definitions

   The terms "reproduce,"    "reproduction," "derivative works,"   and
   "distribution" have  the same meaning here  as under U.S. copyright
   law.

   A  "contribution" is the  original  software, or  any additions  or
   changes to the software.

   A "contributor"  is any  person  that distributes  its contribution
   under this license.

   "Licensed  patents" are  a contributor's  patent  claims that  read
   directly on its contribution.

2. Grant of Rights

   (A) Copyright   Grant-  Subject to  the   terms  of  this  license,
   including the license conditions and limitations in section 3, each
   contributor grants   you a  non-exclusive,  worldwide, royalty-free
   copyright license to reproduce its contribution, prepare derivative
   works of its  contribution, and distribute  its contribution or any
   derivative works that you create.

   (B) Patent Grant-  Subject to the terms  of this license, including
   the   license   conditions and   limitations   in  section  3, each
   contributor grants you   a non-exclusive, worldwide,   royalty-free
   license under its licensed  patents to make,  have made, use, sell,
   offer   for   sale,  import,  and/or   otherwise   dispose  of  its
   contribution   in  the  software   or   derivative  works  of   the
   contribution in the software.

3. Conditions and Limitations

   (A) Reciprocal  Grants- For any  file you distribute  that contains
   code from the software (in source code  or binary format), you must
   provide recipients the source code  to that file  along with a copy
   of this  license,  which license  will  govern that  file.  You may
   license other  files that are  entirely  your own  work and  do not
   contain code from the software under any terms you choose.

   (B) No Trademark License- This license does not grant you rights to
   use any contributors' name, logo, or trademarks.

   (C)  If you  bring  a patent claim    against any contributor  over
   patents that you claim  are infringed by  the software, your patent
   license from such contributor to the software ends automatically.

   (D) If you distribute any portion of the  software, you must retain
   all copyright, patent, trademark,  and attribution notices that are
   present in the software.

   (E) If  you distribute any  portion of the  software in source code
   form, you may do so only under this license by including a complete
   copy of this license with your  distribution. If you distribute any
   portion  of the software in  compiled or object  code form, you may
   only do so under a license that complies with this license.

   (F) The  software is licensed  "as-is." You bear  the risk of using
   it.  The contributors  give no  express  warranties, guarantees  or
   conditions. You   may have additional  consumer  rights  under your
   local laws   which  this license  cannot   change. To   the  extent
   permitted under  your local  laws,   the contributors  exclude  the
   implied warranties of   merchantability, fitness for  a  particular
   purpose and non-infringement.


*****/
#endregion
using System;
using System.Collections.Generic;
using System.Management.Automation;
using System.Management.Automation.Provider;
using System.Text;
using System.Xml;
using System.Runtime.InteropServices;

using Microsoft.Office.Interop.OneNote;

namespace Microsoft.Office.OneNote.PowerShell.Provider
{
    /// <summary>
    /// Implements the OneNote powershell provider. This lets you script your OneNote interactions
    /// in Powershell. This inerits the NavigationCmdletProvider which means all drive capabilities are avaliable. 
    /// </summary>
    [CmdletProvider( "OneNote", ProviderCapabilities.None )]
    public class OneNoteProvider: NavigationCmdletProvider, IContentCmdletProvider
    {
        /// <summary>
        /// Determines if a path is syntactically correct.
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        protected override bool IsValidPath(string path)
        {
            if (String.IsNullOrEmpty(path))
            {
                return false;
            }
            
            //
            //  TODO: There's got to be more than this.
            //
            return true;
        }

        /// <summary>
        /// Determines if there is a OneNote *something* at the path (Notebook, section group, section, or page).
        /// </summary>
        /// <param name="path">A PSPath.</param>
        /// <returns>True if there is an item at this point in the path, false otherwise.</returns>
        protected override bool ItemExists(string path)
        {
            return (getOneNoteNode(path) != null);
        }

        protected override void GetItem(string path)
        {
            if (PathIsDrive(path))
            {
                WriteItemObject(this.PSDriveInfo, path, true);
            }
            XmlNode node = getOneNoteNode(path);
            if (node != null)
            {
                WriteItemObject(node, path, false);
            } else
            {
                WriteError(
                    new ErrorRecord(
                        new ArgumentException("Could not find item"),
                        "InvalidArgument",
                        ErrorCategory.InvalidArgument,
                        path)
                );
            }
        }

        protected override bool HasChildItems(string path)
        {
            XmlNode node = getOneNoteNode(path);
            if (node == null)
            {
                return false;
            }
            return (node.ChildNodes.Count > 0);
        }

        protected override bool IsItemContainer(string path)
        {
            XmlNode node = getOneNoteNode(path);
            return isNodeContainer(node);
        }

        /// <summary>
        /// Determines if a node in the OneNote XML hierarchy represents a container. Containers are 
        /// notebooks, section groups, and sections.
        /// </summary>
        /// <param name="node">The node in question.</param>
        /// <returns>True if the node is a container, false otherwise.</returns>
        private static bool isNodeContainer(XmlNode node)
        {
            if (node == null)
            {
                return false;
            }
            string name = node.LocalName;
            return ((name == "Notebooks") || (name == "Notebook") || (name == "SectionGroup") || (name == "Section"));
        }

        /// <summary>
        /// Gets all items below a particular child.
        /// </summary>
        /// <param name="path"></param>
        /// <param name="recurse"></param>
        protected override void GetChildItems(string path, bool recurse)
        {
            XmlNode node = getOneNoteNode(path);
            if (node == null)
            {
                WriteError(
                    new ErrorRecord(
                    new ArgumentException("Path not valid"),
                    "InvalidArgument",
                    ErrorCategory.InvalidArgument,
                    path
                    )
                );
                return;
            }
            GetChildItems(path, recurse, node);
        }

        /// <summary>
        /// This override of the standard <c>GetChildItems</c> gets to skip the "lookup the XML node
        /// associated with a path" step, because the node is already found. Very useful for recursion.
        /// </summary>
        /// <param name="path">The path to search from.</param>
        /// <param name="recurse">If true, recurse down through the containers.</param>
        /// <param name="node">The OneNote XML node corresponding to <c>path</c>.</param>
        private void GetChildItems(string path, bool recurse, XmlNode node)
        {
            List<KeyValuePair<string, XmlNode>> childContainers = new List<KeyValuePair<string, XmlNode>>( );
            foreach (XmlNode child in node.ChildNodes)
            {
                string childPath = SessionState.Path.Combine(path, getOneNoteNodeName(child));
                bool childContainer = isNodeContainer(child);
                WriteItemObject(child, childPath, childContainer);
                if (childContainer && recurse)
                {
                    childContainers.Add(new KeyValuePair<string, XmlNode>(childPath, child));
                }
            }
            if (recurse)
            {
                foreach (KeyValuePair<string, XmlNode> key in childContainers)
                {
                    GetChildItems(key.Key, true, key.Value);
                }
            }
        }

        protected override void GetChildNames(string path, ReturnContainers returnContainers)
        {
            XmlNode node = getOneNoteNode(path);
            if (node == null)
            {
                WriteInvalidPathError(path);
                return;
            }
            foreach (XmlNode child in node.ChildNodes)
            {
                string childName = getOneNoteNodeName(child);
                string childPath = SessionState.Path.Combine(path, childName);
                bool container = isNodeContainer(child);
                WriteItemObject(childName, childPath, container);
            }
        }

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        /// <summary>
        /// The default action for a OneNote path is to navigate OneNote to that page.
        /// </summary>
        /// <param name="path"></param>
        protected override void InvokeDefaultAction(string path)
        {
            XmlNode node = getOneNoteNode(path);
            if (node == null)
            {
                WriteInvalidPathError(path);
                return;
            }
            Microsoft.Office.Interop.OneNote.ApplicationClass app = getOneNoteApplication( );
            string id = node.Attributes["ID"].Value;
            string name = node.Attributes["name"].Value;
            app.NavigateTo(id, "", true);
            System.Diagnostics.Process[] processes = System.Diagnostics.Process.GetProcesses( );
            string windowTitle = String.Format("{0} - Microsoft Office OneNote", name);
            WriteVerbose("Looking for window: " + windowTitle);
            bool foundWindow = false;
            foreach (System.Diagnostics.Process p in processes)
            {
                //
                //  BUGBUG: This will only work on English.
                //

                if (!String.IsNullOrEmpty(p.MainWindowTitle))
                {
                    WriteDebug("Found window " + p.MainWindowTitle);
                }
                if ((p.MainWindowTitle == windowTitle) && (p.MainWindowHandle != (IntPtr)0))
                {
                    SetForegroundWindow(p.MainWindowHandle);
                    foundWindow = true;
                    break;
                }
            }
            if (!foundWindow)
            {
                //
                //  I'll settle for activating any OneNote window.
                //

                processes = System.Diagnostics.Process.GetProcessesByName("onenote");
                foreach (System.Diagnostics.Process p in processes)
                {
                    if (p.MainWindowHandle != (IntPtr)0)
                    {
                        SetForegroundWindow(p.MainWindowHandle);
                        break;
                    }
                }
            }
        }

        /// <summary>
        /// Helper routine to say that the path was invalid.
        /// </summary>
        /// <param name="path"></param>
        private void WriteInvalidPathError(string path)
        {
            WriteError(new ErrorRecord(new ArgumentException("Path not valid"),
                "InvalidArgument",
                ErrorCategory.InvalidArgument,
                path));
        }

        /// <summary>
        /// Helper routine to get the OneNote application.
        /// </summary>
        /// <returns>A pointer to the OneNote application associated with this provider.</returns>
        private Microsoft.Office.Interop.OneNote.ApplicationClass getOneNoteApplication( )
        {
            return new ApplicationClass( );
        }

        /// <summary>
        /// Clears a OneNote item. This can really only be used against Notebooks, which results
        /// in closing the notebook.
        /// </summary>
        /// <param name="path">Path to the notebook to close.</param>
        protected override void ClearItem(string path)
        {
            XmlNode notebook = getOneNoteNode(path);
            if ((notebook == null) || (notebook.LocalName != "Notebook"))
            {
                WriteInvalidPathError(path);
                return;
            }
            Microsoft.Office.Interop.OneNote.ApplicationClass app = getOneNoteApplication( );
            string id = notebook.Attributes["ID"].Value;
            app.CloseNotebook(id);
        }

        public string[] OneNoteTypes = { "Notebook", "Section", "Page", "Directory", "Group" };

        /// <summary>
        /// Creates a new OneNote "thing" -- notebook, section, section group, or page.
        /// </summary>
        /// <param name="path">The path of the new thing to create.</param>
        /// <param name="itemTypeName">The type of thing to create.</param>
        /// <param name="newItemValue">In the notebook creation case, this parameter is the directory in which
        /// the notebook directory will get created.</param>
        protected override void NewItem(string path, string itemTypeName, object newItemValue)
        {
            WriteDebug("In NewItem. Path = " + path);
            string parentPath = SessionState.Path.ParseParent(path, null);
            string childName = SessionState.Path.ParseChildName(path);
            WriteDebug("Parsed path as: Parent = " + parentPath + ", child = " + childName);
            XmlNode parent = getOneNoteNode(parentPath);
            if (parent == null)
            {
                WriteInvalidPathError(parentPath);
                return;
            }

            //
            //  Find out what we're creating. Per guidelines, itemTypeName is case-insensitive,
            //  and it's sufficient for it to be enough characters to disambiguate between the options
            //  in OneNoteTypes.
            //

            List<string> candidates = new List<string>( );
            Microsoft.Office.OneNote.PowerShell.Utilities.GetCandidatesForString(OneNoteTypes, itemTypeName, candidates);
            if (candidates.Count > 1)
            {
                WriteError(new ErrorRecord(
                    new ArgumentException("itemTypeName is ambiguous"),
                    "InvalidArgument",
                    ErrorCategory.InvalidArgument,
                    itemTypeName)
                );
                return;
            }
            if (candidates.Count == 0)
            {
                WriteError(new ErrorRecord(
                    new ArgumentException("itemTypeName not a recognized OneNote type"),
                    "InvalidArgument",
                    ErrorCategory.InvalidArgument,
                    itemTypeName)
                );
                return;
            }
            string selectedType = candidates[0];
            Microsoft.Office.Interop.OneNote.Application2Class app = new Application2Class();
            string id;
            switch (selectedType)
            {
                case "Notebook":
                    string notebookPath = SessionState.Path.Combine((string)newItemValue, childName);
                    WriteVerbose("Creating notebook in " + notebookPath);
                    app.OpenHierarchy(notebookPath, "", out id, CreateFileType.cftNotebook);
                    break;

                case "Directory":
                case "Section":
                    if (parent.LocalName != "Notebook")
                    {
                        WriteError(new ErrorRecord(new ArgumentException(path + " is not a valid section path: It is not contained in a notebook."),
                            "InvalidArgument",
                            ErrorCategory.InvalidArgument,
                            path));
                        return;
                    }
                    WriteVerbose("Creating section " + childName +
                        " in notebook ID " + parent.Attributes["ID"].Value);
                    app.OpenHierarchy(childName + ".one",
                        parent.Attributes["ID"].Value,
                        out id, 
                        CreateFileType.cftSection);
                    break;

                case "Group":
                    if (parent.LocalName != "Notebook")
                    {
                        WriteError(new ErrorRecord(new ArgumentException(path + " is not a valid section group path: It is not contained in a notebook."),
                            "InvalidArgument",
                            ErrorCategory.InvalidArgument,
                            path));
                        return;
                    }
                    WriteVerbose("Creating Section Group" + childName + " in notebook ID " + parent.Attributes["ID"].Value);
                    app.OpenHierarchy(childName,
                        parent.Attributes["ID"].Value,
                        out id,
                        CreateFileType.cftFolder
                    );
                    break;

                case "Page":

                    //
                    //  Creating pages is a bit more work than creating notebooks or sections. I need to build
                    //  up a bit of XML describing the page creation operation.
                    //

                    WriteVerbose("Creating page " + childName);
                    string parentId = parent.Attributes["ID"].Value;
                    /*
                    //OneNoteXml command = new OneNoteXml();
                    //XmlElement section = command.CreateSection(null, parentId);
                    //command.Document.AppendChild(section);
                    ////OneNoteXml command = new OneNoteXml(parent.OuterXml);
                    ////var section = command.Document.FirstChild;
                    //XmlElement page = command.CreatePage(childName, null);
                    section.AppendChild(page);
                    string xml = command.Document.OuterXml;
                    WriteVerbose(xml);
                    app.UpdateHierarchy(xml, XMLSchema.xs2013);
                    */
                    
                    string pageid;
                    app.CreateNewPage(parentId, out pageid);
                    string xml;
                    app.GetPageContent(pageid, out xml);
                    WriteVerbose("Before:" + xml);
                    var cmd = new OneNoteXml(xml);
                    XmlElement node = (XmlElement)cmd.Document.SelectSingleNode("//one:T", cmd.nsmgr);
                    node.IsEmpty = true;
                    node.AppendChild(cmd.Document.CreateCDataSection(childName));
                    app.UpdatePageContent(cmd.Document.OuterXml, DateTime.MinValue, XMLSchema.xs2013);


                    break;
            }
        }

        /// <summary>
        /// Removes a OneNote item.
        /// </summary>
        /// <param name="path"></param>
        /// <param name="recurse"></param>
        protected override void RemoveItem(string path, bool recurse)
        {
            XmlNode node = getOneNoteNode(path);
            if (node == null)
            {
                WriteInvalidPathError(path);
                return;
            }
            if (ShouldProcess(path))
            {
                if ((node.ChildNodes.Count > 0) && !recurse &&
                    !ShouldContinue(path + " contains children, but -recurse was not specified. Continue?", "Remove")) 
                {
                    return;
                }
                Microsoft.Office.Interop.OneNote.ApplicationClass app = getOneNoteApplication( );
                app.DeleteHierarchy(node.Attributes["ID"].Value, DateTime.MinValue);
            }
        }

        /// <summary>
        /// Gets the OneNote XML node that is associated with a "Path" of names that go:
        /// \Notebook\Section\Page
        /// </summary>
        /// <param name="path">The path to a OneNote organizational element.</param>
        /// <returns>An XML Node representing the element if it's found, or null otherwise.</returns>
        private XmlNode getOneNoteNode(string path)
        {
            //
            //  First, get the current hierarchy down to the page level.
            //

            string xml;
            ApplicationClass _app = new ApplicationClass( );
            _app.GetHierarchy(
                "",
                HierarchyScope.hsPages,
                out xml, XMLSchema.xs2013);
            XmlDocument hierarchy = new XmlDocument( );
            hierarchy.LoadXml(xml);

            if (PathIsDrive(path))
            {
                return hierarchy.DocumentElement;
            }

            //
            //  Next, split the path.
            //

            string[] pathComponents = splitPath(path);

            //
            //  Now, walk down the hierarchy looking at element names.
            //

            XmlNode currentElement = hierarchy.DocumentElement;
            foreach (string component in pathComponents)
            {
                currentElement = getChildNodeByName(currentElement, component);
                if (currentElement == null)
                {
                    return null;
                }
            }
            return currentElement;
        }

        /// <summary>
        /// Copies an item from one location to another.
        /// </summary>
        /// <param name="path"></param>
        /// <param name="copyPath"></param>
        /// <param name="recurse"></param>
        protected override void CopyItem(string path, string copyPath, bool recurse)
        {
            XmlNode source = getOneNoteNode(path);
            XmlNode dest = getOneNoteNode(copyPath);

            //
            //  Validate the containment hierarchy would be preserved after the copy.
            //

            bool validCopy = false;
            switch (source.LocalName)
            {
                case "Page":
                    validCopy = dest.LocalName == "Section";
                    break;
                    
                case "Section":
                case "SectionGroup":
                    validCopy = ((dest.LocalName == "SectionGroup") || (dest.LocalName == "Notebook"));
                    break;

                default:
                    validCopy = false;
                    break;
            }
            if (!validCopy)
            {
                WriteError(new ErrorRecord(new ArgumentException("It is not valid to copy a " + source.LocalName + " to a " + dest.LocalName),
                    "ArgumentException",
                    ErrorCategory.InvalidArgument,
                    source));
                return;
            }
            if ((source.LocalName != "Page") && !recurse)
            {
                if (!ShouldContinue(path + " is a " + source.LocalName + ", but -recurse was not specified. Do a recursive copy?",
                    "Copy child items?"))
                {

                    return;
                }
                recurse = true;
            }
            ApplicationClass _app = new ApplicationClass( );
            copyItems(_app, source, dest, recurse);
        }

        /// <summary>
        /// Internal routine to copy from a source OneNote node to a destination OneNote node.
        /// </summary>
        /// <param name="source"></param>
        /// <param name="dest"></param>
        private void copyItems(ApplicationClass _app, XmlNode source, XmlNode dest, bool recurse)
        {
            if (source.LocalName == "Page")
            {
                string sourceId = source.Attributes["ID"].Value;
                string destId = dest.Attributes["ID"].Value;
                string newPageId;
                string pageContent;
                _app.GetPageContent(sourceId, out pageContent, PageInfo.piAll);
                _app.CreateNewPage(destId, out newPageId, NewPageStyle.npsDefault);
                XmlDocument page = new XmlDocument( );
                page.LoadXml(pageContent);
                page.DocumentElement.Attributes["ID"].Value = newPageId;
                WriteDebug("Just changed PageID to " + newPageId);
                XmlNodeList objectNodes = page.SelectNodes("//*[@objectID]");
                foreach (XmlNode objectNode in objectNodes)
                {
                    XmlAttribute oid = objectNode.Attributes["objectID"];
                    objectNode.Attributes.Remove(oid);
                }
                WriteDebug("Just removed object IDs from " + objectNodes.Count + " node(s)");
                string debugXml = Microsoft.Office.OneNote.PowerShell.Utilities.PrettyPrintXml(page);
                WriteDebug(debugXml.Substring(0, 256));
                _app.UpdatePageContent(page.OuterXml, DateTime.MinValue, XMLSchema.xs2013);
            }
            if (recurse)
            {
                throw new NotImplementedException( );
            }
        }

        /// <summary>
        /// Gets the child node of something in the OneNote hierarchy given its name.
        /// </summary>
        /// <param name="parentElement">The parent node in the OneNote hierarchy.</param>
        /// <param name="component">The name of the child node in the hierarchy.</param>
        /// <returns>A pointer to the child node if found; null otherwise.</returns>
        private XmlNode getChildNodeByName(XmlNode parentElement, string component)
        {
            string childName = unescapeOneNoteName(component);
            WriteDebug("Looking for " + component);
            if (String.Compare(component, "UnfiledNotes", true) == 0)
            {
                //
                //  See if I can get an UnfiledNotes element from the parent.
                //

                WriteDebug("Looking for the special UnfiledNotes element");
                XmlNamespaceManager nsmgr = new XmlNamespaceManager(parentElement.OwnerDocument.NameTable);
                nsmgr.AddNamespace("one", OneNoteXml.OneNoteSchema);
                XmlNode unfiled = parentElement.SelectSingleNode("one:UnfiledNotes",nsmgr);
                if (unfiled != null)
                {
                    WriteDebug("UnfiledNotes element found.");
                    
                    //
                    //  We need to return a pointer to the section contained herein.
                    //  HACK: If the UnfiledNotes notebook ever contains more than one section, this
                    //        will break.
                    //
                    
                    return unfiled.SelectSingleNode("one:Section", nsmgr);
                }
                WriteDebug("UnfiledNotes not found.");
            }
            if (String.Compare(component, "OpenSections", true) == 0)
            {
                WriteDebug("Looking for the special OpenSections element");
                XmlNamespaceManager nsmgr = new XmlNamespaceManager(parentElement.OwnerDocument.NameTable);
                nsmgr.AddNamespace("one", OneNoteXml.OneNoteSchema);
                XmlNode openSections = parentElement.SelectSingleNode("one:OpenSections", nsmgr);
                if (openSections != null)
                {
                    return openSections;
                }
                WriteDebug("OpenSections not found.");
            }
            foreach (XmlNode child in parentElement.ChildNodes)
            {
                XmlAttribute nameAttr = (XmlAttribute)child.Attributes.GetNamedItem("name");
                if (nameAttr == null)
                {
                    continue;
                }
                if (String.Compare(nameAttr.Value, childName, true) == 0)
                {
                    return child;
                }
            }
            return null;
        }

        /// <summary>
        /// This helper routine gets the user-visible name associated with a OneNote node.
        /// </summary>
        /// <param name="node">The OneNote XML node we need to name.</param>
        /// <returns>A string name.</returns>
        private string getOneNoteNodeName(XmlNode node)
        {
            XmlAttribute nameAttr = node.Attributes["name"];
            if (nameAttr != null)
            {
                return escapeOneNoteName(nameAttr.Value);
            }

            //
            //  Hmm. A null name attribute is used for Unfiled Notes and Open Sections.
            //

            return node.LocalName;
        }

        /// <summary>
        /// Given the name of something in OneNote (notebook, section, whatever), this routine escapes the string
        /// so "dangerous" characters (like slashes) don't appear.
        /// </summary>
        /// <param name="name">the name to escape.</param>
        /// <returns>The escaped string.</returns>
        private static string escapeOneNoteName(string name)
        {
            name = name.Replace("/", "&#47;");
            name = name.Replace("\\", "&#92;");
            name = name.Replace(":", "&#58;");
            return name;
        }

        /// <summary>
        /// The compliment to <c>escapeOneNoteName</c>.
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        private static string unescapeOneNoteName(string name)
        {
            name = name.Replace("&#47;", "/");
            name = name.Replace("&#92;", "\\");
            name = name.Replace("&#58;", ":");
            return name;
        }

        #region DriveCmdletProvider

        /// <summary>
        /// Creates a new "drive" representing a connection to OneNote.
        /// </summary>
        /// <param name="drive"></param>
        /// <returns></returns>
        protected override PSDriveInfo NewDrive(PSDriveInfo drive)
        {
            if (drive == null)
            {
                WriteError(new ErrorRecord(
                    new ArgumentNullException("drive"),
                    "NullDrive",
                    ErrorCategory.InvalidArgument,
                    null));
                return null;
            }
            if (String.IsNullOrEmpty(drive.Root))
            {
                WriteError(new ErrorRecord(
                    new ArgumentNullException("drive.Root"),
                    "NoRoot",
                    ErrorCategory.InvalidArgument,
                    null));
                return null;
            }
            OneNoteDriveInfo onDriveInfo = new OneNoteDriveInfo(drive);
            return onDriveInfo;
        }

        #endregion

        #region Path Helper Routines

        private string PathSeparator = "\\";

        /// <summary>
        /// Determines if a path is simply the drive specifier for this provider. This code is copied from the
        /// Windows SDK.
        /// </summary>
        /// <param name="path">The path to test.</param>
        /// <returns>True if the path is just the drive specifier.</returns>
        private bool PathIsDrive(string path)
        {
            if (String.IsNullOrEmpty(path.Replace(this.PSDriveInfo.Root, "")) ||
                String.IsNullOrEmpty(path.Replace(this.PSDriveInfo.Root + PathSeparator, "")))
            {

                return true;
            } else
            {
                return false;
            }
        }

        /// <summary>
        /// Divides up a path based into components based on the path separator. Essentially copied from the
        /// Windows SDK.
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        private string[] splitPath(string path)
        {
            string pathNoDrive = path.Replace(this.PSDriveInfo.Root + PathSeparator, "");
            return pathNoDrive.Split(PathSeparator.ToCharArray( ));
        }
        #endregion

        /// <summary>
        /// Creates the default OneNote: drive for manipulating OneNote notebooks.
        /// </summary>
        /// <returns></returns>
        protected override System.Collections.ObjectModel.Collection<PSDriveInfo> InitializeDefaultDrives( )
        {
            PSDriveInfo di = new PSDriveInfo("OneNote", this.ProviderInfo, "OneNote", "OneNote Notebooks", this.Credential);
            System.Collections.ObjectModel.Collection<PSDriveInfo> drives = base.InitializeDefaultDrives( );
            drives.Add(di);
            return drives;
        }

        #region IContentCmdletProvider Members

        /// <summary>
        /// Removes all of the outline elements in a OneNote page.
        /// </summary>
        /// <param name="path"></param>
        public void ClearContent(string path)
        {
            WriteVerbose("In ClearContent for " + path);
            XmlNode node = getOneNoteNode(path);
            if (node == null)
            {
                WriteInvalidPathError(path);
                return;
            }
            if (node.LocalName != "Page")
            {
                WriteError(new ErrorRecord(new ArgumentException("You may only clear the content of OneNote pages."),
                    "InvalidArgument",
                    ErrorCategory.InvalidArgument,
                    path));
                return;
            }
            WriteVerbose(path + " is a valid OneNote page.");
            string pageXml;
            ApplicationClass app = getOneNoteApplication( );
            string pageId = node.Attributes["ID"].Value;
            WriteVerbose("Page ID is " + pageId);
            app.GetPageContent(pageId, out pageXml, PageInfo.piBasic);
            XmlDocument doc = new XmlDocument( );
            doc.LoadXml(pageXml);
            XmlNamespaceManager nsmgr = new XmlNamespaceManager(doc.NameTable);
            nsmgr.AddNamespace("one", OneNoteXml.OneNoteSchema);
            XmlNodeList outlineElements = doc.SelectNodes("//one:Outline", nsmgr);
            foreach (XmlNode outline in outlineElements)
            {
                string outlineId = outline.Attributes["objectID"].Value;
                WriteVerbose( "Deleting outline: " + outlineId );
                app.DeletePageContent(pageId, outlineId, DateTime.MinValue);
            }
        }

        public object ClearContentDynamicParameters(string path)
        {
            return null;
        }

        public IContentReader GetContentReader(string path)
        {
            XmlNode node = getOneNoteNode(path);
            if (node == null)
            {
                throw new ArgumentException("Path not found", "path");
            }
            ContentReader reader = new ContentReader(
                (OneNoteDriveInfo)PSDriveInfo,
                node
                );
            return reader;
        }

        public object GetContentReaderDynamicParameters(string path)
        {
            return null;
        }

        public IContentWriter GetContentWriter(string path)
        {
            XmlNode node = getOneNoteNode(path);
            if (node == null)
            {
                throw new ArgumentException("Path not found");
            }
            if (node.LocalName != "Page")
            {
                throw new ArgumentException("Path is not a OneNote page");
            }
            return new ContentWriter(node);
        }

        public object GetContentWriterDynamicParameters(string path)
        {
            return null;
        }

        #endregion
    }

    /// <summary>
    /// Maintains the connection to OneNote.
    /// </summary>
    class OneNoteDriveInfo : PSDriveInfo
    {
        /// <summary>
        /// Constructor.
        /// </summary>
        /// <param name="driveInfo"></param>
        public OneNoteDriveInfo(PSDriveInfo driveInfo)
            : base(driveInfo)
        {
        }
    }
}
