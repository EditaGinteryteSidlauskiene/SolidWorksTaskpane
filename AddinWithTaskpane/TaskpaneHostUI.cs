using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Collections.Generic;
using System.Net.Configuration;
using System.Runtime.InteropServices;
using System.Runtime.Remoting.Metadata.W3cXsd2001;
using System.Windows.Forms;
using WarningAndErrorService;

namespace AddinWithTaskpane
{
    //Setting prog id the string in variable SWTASKPANE_PROGID, and then UserControl (below) gets the prog id.
    // User control will be injected into the SolidWorks by passing this id.
    [ProgId(TaskpaneIntegration.SWTASKPANE_PROGID)]
    public partial class TaskpaneHostUI : UserControl
    {
        private WarningService warningService = new WarningService();

        // Getting SW application reference.
        SldWorks swApp;
        public void AssignSWApp(SldWorks inputswApp)
        {
            swApp = inputswApp;
        }
        // Specific assebly
        private ModelDoc2 mainAssembly;

        public TaskpaneHostUI()
        {
            InitializeComponent();
        }

        /// <summary>
        /// Adding functionality to "Open Assembly" button
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button1_Click(object sender, System.EventArgs e)
        {
            // The path of the assembly to be opened
            const string assemblyPath = "C:\\Users\\Edita\\Desktop\\Assembly.sldasm";

            // Open assembly
            mainAssembly = swApp.OpenDoc6(assemblyPath, (int)swDocumentTypes_e.swDocASSEMBLY, (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", 0, 0);

            ModelDoc2 assembly = (ModelDoc2)swApp.ActiveDoc;
        }

        private void button2_Click(object sender, System.EventArgs e)
        {
            DishedEndCollection dishedEndCollection = new DishedEndCollection(warningService, swApp, (ModelDoc2)swApp.ActiveDoc);

            dishedEndCollection.SetNumberOfDishedEnds(0, DishedEndAlignment.Left, 1.2);
        }

        private void button3_Click(object sender, System.EventArgs e)
        {
            DishedEndCollection dishedEndCollection = new DishedEndCollection(warningService, swApp, (ModelDoc2)swApp.ActiveDoc);

            dishedEndCollection.SetNumberOfDishedEnds(2, DishedEndAlignment.Right, 1.2);

        }

        private void button4_Click(object sender, System.EventArgs e)
        {
            ModelDoc2 assemblyModelDoc = (ModelDoc2)swApp.ActiveDoc;
            AssemblyDoc assembly = (AssemblyDoc)swApp.ActiveDoc;
            ModelDocExtension modelDocExtension = assemblyModelDoc.Extension;

            CylindricalShellCollection cylindricalShells = new CylindricalShellCollection(warningService, swApp, assemblyModelDoc);

           
        }
    }
}
