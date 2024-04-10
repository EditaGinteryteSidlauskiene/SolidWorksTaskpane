using SolidWorks.Interop.sldworks;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WarningAndErrorService;

namespace AddinWithTaskpane
{
    internal class CylindricalShellCollection
    {
        SldWorks solidWorksApplication;
        private readonly ModelDoc2 assemblyModelDoc;
        private readonly AssemblyDoc assemblyDocument;

        public CylindricalShell firstCylindricalShell;
        private readonly Collection<CylindricalShell> cylindricalShells = new Collection<CylindricalShell>();

        private readonly WarningService warningService;

        private bool GetCylindricalShells()
        {
            // Stores components
            List<Component2> componentCollection = Utilities.GetTopLevelComponents(assemblyModelDoc);

            // There shall be at least 2 dished end components in the assembly.
            if (componentCollection.Count() < 1)
            {
                warningService.AddError("There shall be at least 1 cylindrical shell component in the assembly.");
                return false;
            }

            //Assign all cylindrical shells
            for (int i = 0; i < componentCollection.Count(); i++)
            {
                CylindricalShell cylindricalShell = new CylindricalShell(warningService, solidWorksApplication, assemblyModelDoc, componentCollection[i]);

                if (warningService.HasErrors())
                {
                    warningService.AddError($"Cylindrical shell {i - 1} object was not created correctly.");
                    return false;
                }

                //Add cylindrical shell into the collection
                cylindricalShells.Add(cylindricalShell);
            }

            return true;
        }

        /// <summary>
        /// Constructor to create a new cylindrical shell object
        /// </summary>
        /// <param name="AssemblyModelDoc"></param>
        public CylindricalShellCollection(WarningService WarningService, SldWorks SolidWorksApplication, ModelDoc2 AssemblyModelDoc)
        {
            warningService = WarningService;

            //If the SW application is null, the method terminates.
            if (SolidWorksApplication is null)
            {
                warningService.AddError("SolidWorksApplication is null in CylindricalShellCollection constructor.");
                return;
            }

            //If the document is null, the method terminates.
            if (AssemblyModelDoc is null)
            {
                warningService.AddError("The AssemblyModelDoc is null in CylindricalShellCollection constructor.");
                return;
            }

            solidWorksApplication = SolidWorksApplication;
            assemblyModelDoc = AssemblyModelDoc;
            assemblyDocument = (AssemblyDoc)assemblyModelDoc;

            GetCylindricalShells();
        }

        /// <summary>
        /// Adds a new cylindrical shell component to the assembly.
        /// </summary>
        /// <param name="DishedEndAlignment">The alignment of the new dished end.</param>
        /// <param name="Distance">The distance of the new dished end from a reference point.</param>
        public void AddCylindricalShell(double Length)
        {
            // Create a new cylindrical shell object with the provided parameter
            CylindricalShell cylindricalShell = new CylindricalShell(
                WarningService: warningService,
                SolidWorksApplication: solidWorksApplication,
                AssemblyModelDoc: assemblyModelDoc,
                Length: Length);

            if (cylindricalShell == null)
            {
                warningService.AddError("CylindricalShell object was not created in AddCylindricalShell method.");
                return;
            }

            // Add the new cylindrical shell to the list
            cylindricalShells.Add(cylindricalShell);
        }

        /// <summary>
        /// Removes the last cylindrical shell component from the assembly.
        /// </summary>
        public void RemoveCylindricalShell()
        {
            // Check if there are any cylindrical shells to remove
            if (cylindricalShells.Count == 0)
            {
                warningService.AddInfo("Cylindrical shell was not removed from RemoveCylindricalShell method as cylindricalShells list is empty.");
                return;
            }

            // Delete the cylindrical shell (assuming it controls the SolidWorks object)
            cylindricalShells.Last().Remove();

            // Remove the last cylindrical shell reference from the tracking list
            cylindricalShells.RemoveAt(cylindricalShells.Count - 1);
        }
    }
}
