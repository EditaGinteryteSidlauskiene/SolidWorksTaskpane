using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Diagnostics;
using System.Linq;
using System.Timers;
using System.Windows.Forms;
using WarningAndErrorService;

namespace AddinWithTaskpane
{
    /// <summary>
    /// Position of dished end
    /// </summary>
    public enum DishedEndAlignment
    { 
        Left,
        Right,
    }

    internal class DishedEnd
    {
        SldWorks solidWorksApplication;
        private ModelDoc2 assemblyModelDoc;
        private AssemblyDoc assemblyDocument;

        //Dished end component's variables
        public Component2 dishedEnd;
        private Feature positionPlane;

        //Mates' variables
        MateCoincident mateCoincidentRightPlane;
        MateCoincident mateCoincidentFrontPlane;
        MateCoincident mateCoincidentAxis;

        //Variable for handling errors
        readonly WarningService warningService;

        private const string COMPONENT_PATH = "C:\\Users\\Edita\\Desktop\\Parts\\MRC 1600x5.SLDPRT";
        private const string COMPARTMENT_DISHED_END_NAME = "Compartment End ";
        private const string RIGHT_PLANE_NAME = "Right Plane";
        private const string FRONT_PLANE_NAME = "Front Plane";
        private const string CENTER_AXIS_NAME = "Center Axis";
        private const string POSITION_PLANE_NAME = "Position Plane";


        /// <summary>
        /// DishedEndPosition property
        /// </summary>
        public DishedEndAlignment DishedEndAlignment
        {
            get => GetAlignment();
            set
            {
                //If the current orientation doesn't match the desired orientation, realign the component
                if (GetAlignment() != value)
                {
                    ChangeAlignment();
                }
            }
        }

        /// <summary>
        /// Represents a reference plane feature within the model.
        /// </summary>
        public Feature ReferencePlane
        {
            //Gets the reference plane feature.
            get
            {
                return positionPlane; 
            }
            
            //Sets the reference plane feature. For internal use only.
            private set 
            { 
                positionPlane = value; 
            } 
        }

        /// <summary>
        /// This constructor is used to assign found dished end and its properties.
        /// </summary>
        /// <param name="WarningService"></param>
        /// <param name="SolidWorksApplication"></param>
        /// <param name="AssemblyModelDoc"></param>
        /// <param name="ExistingDishedEnd"></param>
        public DishedEnd(WarningService WarningService, 
            SldWorks SolidWorksApplication, 
            ModelDoc2 AssemblyModelDoc, 
            Component2 ExistingDishedEnd)
        {
            warningService = WarningService;

            //If the SW application is null, the method terminates.
            if (SolidWorksApplication is null)
            {
                warningService.AddError("SolidWorksApplication is null in InitialAssignment");
                return;
            }

            //If the document is null, the method terminates.
            if (AssemblyModelDoc is null)
            {
                warningService.AddError("The AssemblyModelDoc is null in InitialAssignment");
                return;
            }

            assemblyModelDoc = AssemblyModelDoc;
            assemblyDocument = (AssemblyDoc)assemblyModelDoc;

            solidWorksApplication = SolidWorksApplication;

            //Assign mates, if mates were not assigned, terminate method
            if (!AssignMates(SolidWorksApplication, AssemblyModelDoc, ExistingDishedEnd))
            {
                warningService.AddError("Not all mates are assigned.");
                return;
            }

            //Assign dished end.
            dishedEnd = ExistingDishedEnd;
        }

        /// <summary>
        /// Constructor to create a new dished end object
        /// </summary>
        /// <param name="AssemblyModelDoc"></param>
        public DishedEnd(WarningService WarningService, 
            SldWorks SolidWorksApplication, 
            ModelDoc2 AssemblyModelDoc, 
            DishedEnd ReferenceDishedEnd, 
            DishedEndAlignment DishedEndAlignment, 
            double Distance, 
            int CompartmentNumber)
        {
            warningService = WarningService;

            //If the SW application is null, the method terminates.
            if (SolidWorksApplication is null)
            {
                warningService.AddError("SolidWorksApplication is null in DishedEnd constructor");
                return;
            }

            //If the document is null, the method terminates.
            if (AssemblyModelDoc is null)
            {
                warningService.AddError("The AssemblyModelDoc is null in DishedEnd constructor.");
                return;
            }

            assemblyModelDoc = AssemblyModelDoc;
            assemblyDocument = (AssemblyDoc)assemblyModelDoc;

            solidWorksApplication = SolidWorksApplication;

            mateCoincidentRightPlane = new MateCoincident(this.assemblyModelDoc);
            mateCoincidentFrontPlane = new MateCoincident(this.assemblyModelDoc);
            mateCoincidentAxis = new MateCoincident(this.assemblyModelDoc);

            //Create a new section
            AddDishedEndComponents(ReferenceDishedEnd, DishedEndAlignment, Distance, CompartmentNumber);
        }

        /// <summary>
        /// Assigns mates and position plane. Return true if all 3 mates were assigned, false if not all mates were assigned.
        /// This method is used to assign found dished end and its mates.
        /// </summary>
        /// <param name="SolidWorksApplication"></param>
        /// <param name="AssemblyModelDoc"></param>
        /// <param name="ExistingDishedEnd"></param>
        /// <returns></returns>
        private bool AssignMates(SldWorks SolidWorksApplication, ModelDoc2 AssemblyModelDoc, Component2 ExistingDishedEnd)
        {
            //Get dished end's mates
            object[] dishedEndMates = ExistingDishedEnd.GetMates();

            //If there are less or none mates, the method is terminated.
            if (dishedEndMates is null || dishedEndMates.Count() < 3)
            {
                warningService.AddError("Not enough mates. There must be at least 3 mates.");
                return false;
            }

            //Loop first three mates to assign
            for (int i = 0; i <= 2; i++)
            {
                Entity componentEntity = null;
                Entity assemblyEntity = null;

                //Get and assign assembly and component entities
                if (!GetAssemblyAndComponentEntities(
                    ComponentToCheck: ExistingDishedEnd,
                    ComponentEntity: ref componentEntity,
                    AssemblyEntity: ref assemblyEntity,
                    Mate: (Feature)dishedEndMates[i]))
                {
                    return false;
                }

                //Assign mate of right planes
                if (Utilities.IsEntityMajorPlane((Feature)componentEntity, AssemblyModelDoc, MajorPlane.Right)
                    && Utilities.IsPlaneParallelToMajorPlane(SolidWorksApplication, (Feature)assemblyEntity, MajorPlane.Right))
                {
                    mateCoincidentRightPlane = new MateCoincident(AssemblyModelDoc, (Feature)dishedEndMates[i]);

                    //Assign position plane
                    positionPlane = (Feature)assemblyEntity;

                    //Go to next mate
                    continue;
                }

                //Assign mate of front planes
                if (Utilities.IsEntityMajorPlane((Feature)assemblyEntity, AssemblyModelDoc, MajorPlane.Front)
                    && Utilities.IsEntityMajorPlane((Feature)componentEntity, AssemblyModelDoc, MajorPlane.Front))
                {
                    mateCoincidentFrontPlane = new MateCoincident(AssemblyModelDoc, (Feature)dishedEndMates[i]);

                    //Go to next mate
                    continue;
                }

                //Assign mate of center axises
                if (Utilities.IsEntityAxis(componentEntity)
                    && Utilities.IsEntityAxis(assemblyEntity))
                {
                    mateCoincidentAxis = new MateCoincident(AssemblyModelDoc, (Feature)dishedEndMates[i]);

                    //Go to next mate
                    continue;
                }
            }

            //Returns true of all mate were assigned, false if not all mates were assigned.
            return
                mateCoincidentRightPlane != null
                && mateCoincidentFrontPlane != null
                && mateCoincidentAxis != null;
        }

        /// <summary>
        /// Attempts to extract assembly and component entities from a set of mate entities, 
        /// specifically in the context of an existing dished end component.
        /// </summary>
        /// <param name="ComponentToCheck">//The existing dished end component.</param>----------------------------------------------
        /// <param name="ComponentEntity">Output: The component entity if found.</param>
        /// <param name="AssemblyEntity">Output: The assembly entity if found.</param>
        /// <param name="mateEntities">An array of mate entities to analyze.</param>
        /// <returns>True if both component and assembly entities were successfully identified, false otherwise.</returns>
        private bool GetAssemblyAndComponentEntities(
            Component2 ComponentToCheck,
            ref Entity ComponentEntity,
            ref Entity AssemblyEntity,
            Feature Mate)
        {
            // Cast the MateFeatureData object to a CoincidentMateFeatureData object. 
            CoincidentMateFeatureData coincMateData = (CoincidentMateFeatureData)Mate.GetDefinition();

            //Get mate's entities
            object[] mateEntities = coincMateData.EntitiesToMate;

            // Validate input:
            if (mateEntities.Length != 2)
            {
                warningService.AddError("Incorrect number of mate entities.");
                return false;
            }

            //Check if the first entity belongs to the component AND the second entity belongs to the assembly
            if (Utilities.IsEntityInComponent((Entity)mateEntities[0])
                && Utilities.IsEntityInSpecificComponent((Entity)mateEntities[0], ComponentToCheck)
                && !Utilities.IsEntityInComponent((Entity)mateEntities[1]))
            {
                // Success: First entity is the component entity, second is the assembly entity 
                ComponentEntity = (Entity)mateEntities[0];
                AssemblyEntity = (Entity)mateEntities[1];
                return true;
            }

            //Check if the second entity belongs to the component AND the first entity belongs to the assembly
            if (Utilities.IsEntityInComponent((Entity)mateEntities[1])
                && Utilities.IsEntityInSpecificComponent((Entity)mateEntities[1], ComponentToCheck)
                && !Utilities.IsEntityInComponent((Entity)mateEntities[0]))
            {
                //Success: Second entity is the component entity, first is the assembly entity 
                ComponentEntity = (Entity)mateEntities[1];
                AssemblyEntity = (Entity)mateEntities[0];
                return true;
            }

            //Find specific errors

            //Error 1: Both entities are in assembly
            if (!Utilities.IsEntityInComponent((Entity)mateEntities[0])
                && !Utilities.IsEntityInComponent((Entity)mateEntities[1]))
            {
                warningService.AddError("Both entities are in the assembly.");
            }

            //Error 2: Both entities are in component
            if (Utilities.IsEntityInComponent((Entity)mateEntities[0])
               && Utilities.IsEntityInComponent((Entity)mateEntities[1]))
            {
                warningService.AddError("Both entities are in the component.");
            }

            //Error 3: One of the entities is not in a specific component.
            if (!Utilities.IsEntityInSpecificComponent((Entity)mateEntities[0], ComponentToCheck)
               || !Utilities.IsEntityInSpecificComponent((Entity)mateEntities[1], ComponentToCheck))
            {
                warningService.AddError("One of entities is in a different component.");
            }

            return false;
        }

        /// <summary>
        /// Gets whether dished end is alligned left or right.
        /// </summary>
        /// <returns></returns>
        private DishedEndAlignment GetAlignment()
        {
            //Get the transformation matrix from the dishedEnd object
            MathTransform transform = dishedEnd.Transform2;

            //Transform the reference point (1, 0, 0) using the transformation matrix
            double[] TransformedVector = Utilities.TransformVector(solidWorksApplication, transform, new double[3] { 1, 0, 0 });

            //Determine orientation based on the transformed point's X-coordinate
            return (TransformedVector[0] > 0 ? DishedEndAlignment.Left : DishedEndAlignment.Right);
        }

        /// <summary>
        /// Adds a dished end and mates it with already existing dished end, assembly's front plane and axis
        /// </summary>
        /// <param name="ReferenceDishedEnd"></param>
        /// <param name="DishedEndAlignment"></param>
        /// <param name="Distance"></param>
        public void AddDishedEndComponents(DishedEnd ReferenceDishedEnd, DishedEndAlignment DishedEndAlignment, double Distance, int CompartmentNumber)
        {
            //Create position plane
            string positionPlaneName = COMPARTMENT_DISHED_END_NAME + (CompartmentNumber) + " " + POSITION_PLANE_NAME;
            positionPlane = Utilities.CreateReferencePlaneWithDistance(assemblyModelDoc, ReferenceDishedEnd.ReferencePlane, Distance, positionPlaneName);

            //Adds a dished end component
            dishedEnd = Utilities.AddComponent(solidWorksApplication, assemblyModelDoc, COMPONENT_PATH);

            //If the component is not added, the method is terminated
            if (dishedEnd is null)
            {
                warningService.AddError("The component's document was not found due to incorrect path.");
                return;
            }

            //Get planes and axis needed to mate

            Feature assemblyCenterAxis = Utilities.GetNthFeatureOfType(assemblyModelDoc, FeatureType.RefAxis, 1);

            //If the given name of axis is incorrect, it cannot be get and the method terminates
            if (assemblyCenterAxis is null)
            {
                warningService.AddError("Component cannot be mated, because there is no axis with a given name");
                return;
            }

            //Rename the component
            Feature componentToRename = Utilities.GetFeatureByName(assemblyModelDoc, dishedEnd.Name2);
            componentToRename.Name = COMPARTMENT_DISHED_END_NAME + (CompartmentNumber);

            //Mate dished end's and assembly's right and front planes, and center axises.
            mateCoincidentRightPlane.CreateMate(
                ComponentFeature1: positionPlane,
                ComponentFeature2: Utilities.GetMajorPlane(dishedEnd, MajorPlane.Right),
                AlignmentType: MateAlignment.Aligned,
                Name: COMPARTMENT_DISHED_END_NAME + (CompartmentNumber) + " - " + RIGHT_PLANE_NAME);

            mateCoincidentFrontPlane.CreateMate(
                ComponentFeature1: Utilities.GetMajorPlane(assemblyModelDoc, MajorPlane.Front),
                ComponentFeature2: Utilities.GetMajorPlane(dishedEnd, MajorPlane.Front),
                AlignmentType: MateAlignment.Aligned,
                Name: COMPARTMENT_DISHED_END_NAME + (CompartmentNumber) + " - " + FRONT_PLANE_NAME);

            mateCoincidentAxis.CreateMate(
                ComponentFeature1: assemblyCenterAxis,
                ComponentFeature2: Utilities.GetNthFeatureOfType(dishedEnd, FeatureType.RefAxis, 1),
                AlignmentType: MateAlignment.Anti_Aligned,
                Name: COMPARTMENT_DISHED_END_NAME + (CompartmentNumber) + " - " + CENTER_AXIS_NAME);

            //Set the required position to the DishedEndPosition property
            this.DishedEndAlignment = DishedEndAlignment;
        }

        /// <summary>
        /// Changes alignment of the dished end component
        /// </summary>
        public void ChangeAlignment()
        {
            //Suppress axis mate
            mateCoincidentAxis.Suppress();

            //Change alignment of the component
            //Warning message if ChangeAlignement() did not work
            if (mateCoincidentFrontPlane.ChangeAlignment() == false)
            {
                warningService.AddWarning("The alignment of the component could not be changed. Check ModifyDefinition() method in MateCoincident.ChangeAlignment()");
            }

            //Change alignment of axis
            //Warning message if ChangeAlignement() did not work
            if (mateCoincidentAxis.ChangeAlignment() == false)
            {
                warningService.AddWarning("The alignment of the axis mate could not be changed. Check ModifyDefinition() method in MateCoincident.ChangeAlignment()");
            }

            //Unsuppress axis mate
            mateCoincidentAxis.Unsuppress();
        }

        /// <summary>
        /// Changes assembly's reference plane, which the component is mated with
        /// </summary>
        /// <param name="NewReferencePlane"></param>
        public void ChangeReferenceEnd(DishedEnd NewReferenceDishedEnd)
        {
            //Get dished end's reference plane
            Feature dishedEndPlane = NewReferenceDishedEnd.ReferencePlane;

            //Change refence plane
            bool status = Utilities.ChangeReferenceOfReferencePlane(assemblyModelDoc, dishedEndPlane, positionPlane);

            //Warning message if ChangeReferenceOfReferencePlane() did not work
            if (status == false)
            {
                warningService.AddWarning("The reference could not be changed. Check ModifyDefinition() method in SolidWorksUtilities.ChangeReferenceOfReferencePlane()");
            }
        }

        /// <summary>
        /// Changes distance of the dished end from the starting plane
        /// </summary>
        /// <param name="Distance"></param>
        public void ChangeDistance(double Distance)
        {
            bool status = Utilities.ChangeDistanceOfReferencePlane(assemblyModelDoc, positionPlane, Distance);

            //Warning message if ChangeDistanceOfReferencePlane() did not work
            if (status == false) warningService.AddWarning("The distande could not be changed. Check ModifyDefinition() method in SolidWorksUtilities.ChangeDistanceOfReferencePlane()");
        }

        /// <summary>
        /// Delete the dished end component
        /// </summary>
        public void Delete()
        {
            SelectionMgr selectionManager = (SelectionMgr)assemblyModelDoc.SelectionManager;
            SelectData selectData = selectionManager.CreateSelectData();

            //Select the dished end to be deleted
            dishedEnd.Select4(false, selectData, false);
            positionPlane.Select2(true, 1);

            //Delete selected dished end
            assemblyDocument.DeleteSelections(0);
        }

        public void Supress()
        {
            Utilities.SupressComponent(dishedEnd);
            Utilities.SupressFeature(this.ReferencePlane);
        }

        public void Unsupress()
        {
            Utilities.UnsupressFeature(this.ReferencePlane);
            Utilities.UnsupressComponent(dishedEnd);
        }
    }
}
