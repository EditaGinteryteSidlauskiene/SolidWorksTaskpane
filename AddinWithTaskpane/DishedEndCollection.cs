using SolidWorks.Interop.sldworks;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using WarningAndErrorService;

namespace AddinWithTaskpane
{
    internal class DishedEndCollection
    {
        SldWorks solidWorksApplication;
        private readonly ModelDoc2 assemblyModelDoc;
        private readonly AssemblyDoc assemblyDocument;

        public DishedEnd leftDishedEnd;
        public DishedEnd rightDishedEnd;
        private readonly Collection<DishedEnd> compartmentDishedEnds = new Collection<DishedEnd>();

        private readonly WarningService warningService;

        private bool GetDishedEnds()
        {
            DishedEnd temporaryObject;

            // Stores components
            List<Component2> componentCollection = Utilities.GetTopLevelComponents(assemblyModelDoc);

            // There shall be at least 2 dished end components in the assembly.
            if (componentCollection.Count() < 2)
            {
                warningService.AddError("There shall be at least 2 dished end components in the assembly.");
                return false;
            }

            //Pass the first component to create a new dished end object and assign it to temporary variable.
            temporaryObject = new DishedEnd(warningService, solidWorksApplication, assemblyModelDoc, componentCollection[0]);

            if (warningService.HasErrors())
            {
                warningService.AddError("Left dished end object was not created correctly.");
                return false;
            }

            //If the dishedEnd is left aligned, it means that this dished end is the left dished end
            //and our temporaryObject variable is assigned to leftDishedEnd variable.
            if (temporaryObject.DishedEndAlignment == DishedEndAlignment.Left)
            {
                leftDishedEnd = temporaryObject;
            }
            else
            {
                return false;
            }

            //Pass the Second component to create a new dished end object and assign it to temporary variable.
            temporaryObject = new DishedEnd(warningService, solidWorksApplication, assemblyModelDoc, componentCollection[1]);

            if (warningService.HasErrors())
            {
                warningService.AddError("Right dished end object was not created correctly.");
                return false;
            }

            //If the dishedEnd is right aligned, it means that this dished end is the right dished end
            //and our temporaryObject variable is assigned to rightDishedEnd variable.
            if (temporaryObject.DishedEndAlignment == DishedEndAlignment.Right)
            {
                rightDishedEnd = temporaryObject;
            }
            else
            {
                return false;
            }

            //Assign remaining dished ends
            for (int i = 2; i < componentCollection.Count(); i++)
            {
                DishedEnd dishedEndCompartment = new DishedEnd(warningService, solidWorksApplication, assemblyModelDoc, componentCollection[i]);
                if (warningService.HasErrors())
                {
                    warningService.AddError($"Compartment end {i-1} object was not created correctly.");
                    return false;
                }
                compartmentDishedEnds.Add(dishedEndCompartment);
            }

            return true;
        }

        /// <summary>
        /// Constructor to create a new dished end object
        /// </summary>
        /// <param name="AssemblyModelDoc"></param>
        public DishedEndCollection(WarningService WarningService, SldWorks SolidWorksApplication, ModelDoc2 AssemblyModelDoc)
        {
            warningService = WarningService;

            //If the SW application is null, the method terminates.
            if (SolidWorksApplication is null)
            {
                warningService.AddError("SolidWorksApplication is null in DishedEndCollection constructor.");
                return;
            }

            //If the document is null, the method terminates.
            if (AssemblyModelDoc is null)
            {
                warningService.AddError("The AssemblyModelDoc is null in DishedEndCollection constructor.");
                return;
            }

            solidWorksApplication = SolidWorksApplication;
            assemblyModelDoc = AssemblyModelDoc;
            assemblyDocument = (AssemblyDoc)assemblyModelDoc;

            GetDishedEnds();
        }

        /// <summary>
        /// Adds a new dished end component to the assembly.
        /// </summary>
        /// <param name="DishedEndAlignment">The alignment of the new dished end.</param>
        /// <param name="Distance">The distance of the new dished end from a reference point.</param>
        private void AddCompartmentEnd(DishedEndAlignment DishedEndAlignment, double Distance)
        {
            // Create a new dished end object with the provided parameters
            DishedEnd compartmentEnd = new DishedEnd(
                WarningService: warningService,
                SolidWorksApplication: solidWorksApplication,
                AssemblyModelDoc: assemblyModelDoc,
                ReferenceDishedEnd: (compartmentDishedEnds.Count == 0 ? leftDishedEnd : compartmentDishedEnds.Last()),
                DishedEndAlignment: DishedEndAlignment,
                Distance: Distance,
                CompartmentNumber: compartmentDishedEnds.Count+1);

            if (compartmentEnd == null)
            {
                warningService.AddError("DishedEnd object was not created in AddCompartmentEnd method.");
                return;
            }

            // Add the new dished end to the list
            compartmentDishedEnds.Add(compartmentEnd);
        }

        /// <summary>
        /// Removes the last dished end component from the assembly.
        /// </summary>
        private void RemoveCompartmentEnd()
        {
            // Check if there are any dished ends to remove
            if (compartmentDishedEnds.Count == 0)
            {
                warningService.AddInfo("Compartment dised end was not removed from RemoveCompartmentEnd method as compartmentDishedEnds list is empty.");
                return;
            }

            // Delete the last dished end (assuming it controls the SolidWorks object)
            compartmentDishedEnds.Last().Delete();

            // Remove the last dished end reference from the tracking list
            compartmentDishedEnds.RemoveAt(compartmentDishedEnds.Count - 1);
        }

        /// <summary>
        /// Manages the number of dished ends to match a required count.
        /// </summary>
        /// <param name="RequiredNumberOfDishedEnds">The desired number of dished ends.</param>
        /// <param name="DefaultDishedEndAlignment">The default alignment to use when adding new dished ends.</param>
        /// <param name="DefaultDistance">The default distance to use when adding new dished ends.</param>
        public void SetNumberOfDishedEnds(int RequiredNumberOfDishedEnds, DishedEndAlignment DefaultDishedEndAlignment, double DefaultDistance)
        {
            // Early exit if the required number already matches the current number
            if (RequiredNumberOfDishedEnds == compartmentDishedEnds.Count) return;

            if (RequiredNumberOfDishedEnds < compartmentDishedEnds.Count)
            {
                // Adjust the reference for the rightmost dished end
                rightDishedEnd.ChangeReferenceEnd(RequiredNumberOfDishedEnds == 0 ? leftDishedEnd : compartmentDishedEnds[RequiredNumberOfDishedEnds - 1]);

                // Remove extra dished ends until the required count is reached
                while (RequiredNumberOfDishedEnds != compartmentDishedEnds.Count)
                {
                    RemoveCompartmentEnd();
                    if (warningService.HasErrors()) return;
                }
            }
            else
            {
                // Add new dished ends until the required count is reached
                while (RequiredNumberOfDishedEnds != compartmentDishedEnds.Count)
                {
                    AddCompartmentEnd(DefaultDishedEndAlignment, DefaultDistance);
                    if (warningService.HasErrors()) return;
                }

                // Adjust the reference for the rightmost dished end
                rightDishedEnd.ChangeReferenceEnd(compartmentDishedEnds.Count == 0 ? leftDishedEnd : compartmentDishedEnds.Last());
            }
        }
    }
}
