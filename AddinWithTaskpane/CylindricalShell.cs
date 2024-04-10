using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using SolidWorks.Interop.swdocumentmgr;
using System;
using System.CodeDom;
using System.Collections.Generic;
using System.Drawing.Text;
using System.Linq;
using System.Net.Configuration;
using System.Security.Cryptography.X509Certificates;
using System.Windows.Forms;
using WarningAndErrorService;

namespace AddinWithTaskpane
{
    internal class CylindricalShell
    {
        //
        // Tikslas:
        // Valdyti assembly su cilindrinem dalim. Cil.Dalys(Partai) gali buti ivairiu ilgiu.
        // Partai turi buti virtualus ir nepriklausomi vienas nuo kito.
        // assemblyje, Feature tree, Componentas su desiniu mygtuku -> Make independent
        // Komponentas buvo virtualus.
        // Parto savybes: 2x reference plane: Right Reference Plane ir Left Reference plane.
        // Metodas: ChangeLength(double) kiecia ilgi.
        // Cilindrine dalis visada bus bent 1. Todel galima priimti, kad asemblyje jau bus bent 1 cilindrine dalis.
        // Darbas turi buti toks: Paimi cilindrine dali is asemblio, padarai jos kopija, padarai nepriklausoma
        // ir sumeitini.
        // Asemblis turi buti panasus kaip dugnu: AddCylindricalPart, RemoveCylindricalPart
        // Cil. dalys asemblyje turi buti sumeitinti 45 laipsniu kampu nuo front Plane del siuliu.

        SldWorks solidWorksApplication;
        private ModelDoc2 assemblyModelDoc;
        private AssemblyDoc assemblyDocument;

        //Cylinder component's variables
        public Component2 cylindricalShell;
        public Feature rightReferencePlane;
        public Feature leftReferencePlane;

        //Mates' variables
        MateCoincident mateCoincidentLeftPlane;
        MateCoincident mateCoincidentAxis;
        public MateAngle mateAngleFrontPlane;

        //Variable for handling errors
        readonly WarningService warningService;

        private const string LEFT_REFERENCE_PLANE_NAME = "Left Plane";
        private const string FRONT_REFERENCE_PLANE_NAME = "Front Plane";
        private const string CENTER_AXIS_NAME = "Center Axis";
        private const string CYLINDRICAL_SHELL_NAME = "Cylindrical Shell";
        private const string CYLINDRICAL_SHELL_PATH = "C:\\Users\\Edita\\Desktop\\Parts\\Shell Cylyndrical ø1600×6 L1000.SLDPRT";

        /// <summary>
        /// This constructor is used to assign found cylindrical shells and their properties.
        /// </summary>
        /// <param name="WarningService"></param>
        /// <param name="SolidWorksApplication"></param>
        /// <param name="AssemblyModelDoc"></param>
        /// <param name="ExistingDishedEnd"></param>
        public CylindricalShell(WarningService WarningService,
            SldWorks SolidWorksApplication,
            ModelDoc2 AssemblyModelDoc,
            Component2 CylindricalShell)
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
            if (!AssignMates(SolidWorksApplication, AssemblyModelDoc, CylindricalShell))
            {
                warningService.AddError("Not all mates were assigned.");
                return;
            }

            //Assign cylindrical shell's right plane
            rightReferencePlane = Utilities.GetNthFeatureOfType(CylindricalShell, FeatureType.RefPlane, 5);

            //Assign cylindrical shell.
            cylindricalShell = CylindricalShell;
        }

        /// <summary>
        /// Constructor creates adds a new cylindrical shell and mates it with the last one.
        /// </summary>
        /// <param name="WarningService"></param>
        /// <param name="SolidWorksApplication"></param>
        /// <param name="AssemblyModelDoc"></param>
        public CylindricalShell(WarningService WarningService,
            SldWorks SolidWorksApplication,
            ModelDoc2 AssemblyModelDoc,
            double Length)
        {
            warningService = WarningService;

            //If the SW application is null, the method terminates.
            if (SolidWorksApplication is null)
            {
                warningService.AddError("SolidWorksApplication is null in CylindricalShell constructor");
                return;
            }

            //If the document is null, the method terminates.
            if (AssemblyModelDoc is null)
            {
                warningService.AddError("The AssemblyModelDoc is null in CylindricalShell constructor.");
                return;
            }

            assemblyModelDoc = AssemblyModelDoc;
            assemblyDocument = (AssemblyDoc)assemblyModelDoc;

            solidWorksApplication = SolidWorksApplication;

            mateCoincidentLeftPlane = new MateCoincident(this.assemblyModelDoc);
            mateCoincidentAxis = new MateCoincident(this.assemblyModelDoc);
            mateAngleFrontPlane = new MateAngle(this.assemblyModelDoc);

            //Adds a cylindrical shell
            AddCylindricalShell();

            //Change length of the cylindrical shell
            ChangeLength(Length);
        }

        /// <summary>
        /// Assigns mates and left reference plane. Return true if all 3 mates were assigned, false if not all mates were assigned.
        /// This method is used to assign found cylindrical shell and its mates.
        /// </summary>
        /// <param name="SolidWorksApplication"></param>
        /// <param name="AssemblyModelDoc"></param>
        /// <param name="CylindricalShell"></param>
        /// <returns></returns>
        private bool AssignMates(SldWorks SolidWorksApplication, ModelDoc2 AssemblyModelDoc, Component2 CylindricalShell)
        {
            //Get cylindrical shelll's mates
            object[] cylindricalShellMates = CylindricalShell.GetMates();

            //If there are less or none mates, the method is terminated.
            if (cylindricalShellMates is null || cylindricalShellMates.Count() < 3)
            {
                warningService.AddError("Not enough mates. There must be at least 3 mates.");
                return false;
            }

            //Loop first three mates to assign
            for (int i = 0; i <= 2; i++)
            {
                Entity currentCylindricShellLeftPlane = null;
                Entity lastPreviousCylindricShellRightPlane = null;
                Entity assemblyReferencePLane = null;
                Entity cylinderShellCenterAxis = null;
                Entity assemblyCenterAxis = null;
                Entity cylindricShellFrontPlane = null;
                Entity assemblyFrontPlane = null;

                //Get and assign assembly and component entities
                if (!GetEntities(
                    ComponentToCheck: CylindricalShell,
                    CurrentComponentLeftPlane: ref currentCylindricShellLeftPlane,
                    PreviousComponentRightPlane: ref lastPreviousCylindricShellRightPlane,
                    AssemblyReferencePlane: ref assemblyReferencePLane,
                    CylindricalShellCenterAxis: ref cylinderShellCenterAxis,
                    AssemblyCenterAxis: ref assemblyCenterAxis,
                    CylindricalShellFrontPlane: ref cylindricShellFrontPlane,
                    AssemblyFrontPlane: ref assemblyFrontPlane,
                    Mate: (Feature)cylindricalShellMates[i]))
                {
                    return false;
                }

                //Assign mate of cylinder shell left plane and assembly's or other cylinder shell's right plane
                if (currentCylindricShellLeftPlane != null &&
                    (lastPreviousCylindricShellRightPlane != null
                    ||
                    assemblyReferencePLane != null))
                {
                    mateCoincidentLeftPlane = new MateCoincident(assemblyModelDoc, (Feature)cylindricalShellMates[i]);

                    //Assing cylinder shell's left reference plane
                    leftReferencePlane = (Feature)currentCylindricShellLeftPlane;

                    continue;
                }

                //Assign mate between cylinder shell's and assembly's axises. 
                if (cylinderShellCenterAxis != null &&
                    assemblyCenterAxis != null)
                {
                    mateCoincidentAxis = new MateCoincident(assemblyModelDoc, (Feature)cylindricalShellMates[i]);

                    continue;
                }

                //Assign angle mate between cylinder shell's and assembly's front planes.
                if (cylindricShellFrontPlane != null &&
                    assemblyFrontPlane != null)
                {
                    mateAngleFrontPlane = new MateAngle(assemblyModelDoc, (Feature)cylindricalShellMates[i]);

                    continue;
                }
            }

            //Returns true if all mate were assigned, false if not all mates were assigned.
            return
                mateCoincidentLeftPlane != null
                && mateCoincidentAxis != null
                && mateAngleFrontPlane != null;
        }

        /// <summary>
        /// Attempts to extract assembly and component entities from a set of mate entities, 
        /// specifically in the context of an existing cylindrical shell component.
        /// </summary>
        /// <param name="ComponentToCheck"></param>
        /// <param name="CurrentComponentLeftPlane"></param>
        /// <param name="PreviousComponentRightPlane"></param>
        /// <param name="AssemblyReferencePlane"></param>
        /// <param name="CylindricalShellCenterAxis"></param>
        /// <param name="AssemblyCenterAxis"></param>
        /// <param name="CylindricalShellFrontPlane"></param>
        /// <param name="AssemblyFrontPlane"></param>
        /// <param name="Mate"></param>
        /// <returns></returns>
        private bool GetEntities(
            Component2 ComponentToCheck,
            ref Entity CurrentComponentLeftPlane,
            ref Entity PreviousComponentRightPlane,
            ref Entity AssemblyReferencePlane,
            ref Entity CylindricalShellCenterAxis,
            ref Entity AssemblyCenterAxis,
            ref Entity CylindricalShellFrontPlane,
            ref Entity AssemblyFrontPlane,
            Feature Mate)
        {
            //Create mateFeature Data to get access to mate's feature data. 
            MateFeatureData mateFeatureData = (MateFeatureData)Mate.GetDefinition();

            //Check if the mate's type is angle
            if (mateFeatureData.TypeName == (int)swMateType_e.swMateANGLE)
            {
                //Cast the MateFeatureData object to a AngleMateFeatureData object.
                AngleMateFeatureData angleMateData = (AngleMateFeatureData)mateFeatureData;

                //Get entities of angle mate
                object[] angleMateEntities = angleMateData.EntitiesToMate;

                // Validate input:
                if (angleMateEntities.Length != 2)
                {
                    warningService.AddError("Incorrect number of mate entities.");
                    return false;
                }

                //If first entity is in a specific compoenent AND second entity is in assembly and this entity is front major plane
                //then this mate is the angle mate of the cylindric shell whose front major plane is mated with assembly's front major plane
                if ((Utilities.IsEntityInComponent((Entity)angleMateEntities[0]) &&                            //Is the first entity in a component?
                    Utilities.IsEntityInSpecificComponent((Entity)angleMateEntities[0], ComponentToCheck))      //Is the first entity in the specific component?
                    &&
                    (!Utilities.IsEntityInComponent((Entity)angleMateEntities[1]) &&                            //Is the second entity in the assembly?
                    Utilities.IsEntityMajorPlane((Feature)angleMateEntities[1], assemblyModelDoc, MajorPlane.Front)))   //Is the second entity the front major plane?
                {
                    CylindricalShellFrontPlane = (Entity)angleMateEntities[0];
                    AssemblyFrontPlane = (Entity)angleMateEntities[1];
                }

                //If first entity is in assembly and it is a front major plane AND second entity is in the specific component
                //then this mate is the angle mate of the cylindric shell whose front major plane is mated with assembly's front major plane
                if ((!Utilities.IsEntityInComponent((Entity)angleMateEntities[0]) &&                            //Is the first entity in the assembly?
                    Utilities.IsEntityMajorPlane((Feature)angleMateEntities[0], assemblyModelDoc, MajorPlane.Front))    //Is the first entity the front major plane?
                    &&
                    (Utilities.IsEntityInComponent((Entity)angleMateEntities[1]) &&                                     //Is the second entity in a component?
                    Utilities.IsEntityInSpecificComponent((Entity)angleMateEntities[1], ComponentToCheck)))        //Is the second entity in the specific component?
                {
                    AssemblyFrontPlane = (Entity)angleMateEntities[0];
                    CylindricalShellFrontPlane = (Entity)angleMateEntities[1];
                }
                return true;
            }

            //Cast the MateFeatureData object to a CoincidentMateFeatureData object.
            CoincidentMateFeatureData coincMateData = (CoincidentMateFeatureData)mateFeatureData;

            //Get mate's entities
            object[] mateEntities = coincMateData.EntitiesToMate;

            // Validate input:
            if (mateEntities.Length != 2)
            {
                warningService.AddError("Incorrect number of mate entities.");
                return false;
            }

            //If first entity is in a specific compoenent AND second entity is in assembly and this entity is parallel to major right plane
            //then this mate is the mate of the cylindric shell whose left reference plane is mated with assembly's reference plane
            if ((Utilities.IsEntityInComponent((Entity)mateEntities[0]) &&                          //Is the first entity in a component?
                Utilities.IsEntityInSpecificComponent((Entity)mateEntities[0], ComponentToCheck))   //Is the first entity in the specific cylindrical shell?
                &&
                (!Utilities.IsEntityInComponent((Entity)mateEntities[1]) &&                         //Is the second entity in assembly?
                Utilities.IsEntityPlane((Entity)mateEntities[1]) &&                                 //Is the second entity a plane?
                Utilities.IsPlaneParallelToMajorPlane(solidWorksApplication, (Feature)mateEntities[1], MajorPlane.Right)))  //Is the second entity parallel to assembly's major plane?
            {
                CurrentComponentLeftPlane = (Entity)mateEntities[0];
                AssemblyReferencePlane = (Entity)mateEntities[1];
                return true;
            }

            //If first entity is in the assembly and this entity is parallel to major right plane AND second entity is in the specific compoenent
            //then this mate is the mate of the cylindric shell whose left reference plane is mated with assembly's reference plane
            if ((!Utilities.IsEntityInComponent((Entity)mateEntities[0]) &&                                     //Is the first entity in assembly?
                Utilities.IsEntityPlane((Entity)mateEntities[0]) &&                                             //Is the first entity a plane?
                Utilities.IsPlaneParallelToMajorPlane(solidWorksApplication, (Feature)mateEntities[0], MajorPlane.Right))       //Is the first entity parallel to assembly's major plane?
                &&
                (Utilities.IsEntityInComponent((Entity)mateEntities[1]) &&                              //Is the second entity in a component?
                Utilities.IsEntityInSpecificComponent((Entity)mateEntities[1], ComponentToCheck)))      //Is the second entity in the specific cylindrical shell?
            {
                AssemblyReferencePlane = (Entity)mateEntities[0];
                CurrentComponentLeftPlane = (Entity)mateEntities[1];
                return true;
            }

            //Get last previous cylindrical shell
            Component2 lastCylindricalShell = Utilities.GetTopLevelComponents(assemblyModelDoc)[0];

            //If first entity is in the specific cylindrical shell and the second entity is in the last previous cylindrical shell
            //then this mate is the mate of the cylindric shell whose left reference plane is mated with pevious cylindracl shell's right reference plane
            if (Utilities.IsEntityInComponent((Entity)mateEntities[0]) &&                          //Is the first entity in a component?
                Utilities.IsEntityInSpecificComponent((Entity)mateEntities[0], ComponentToCheck)    //Is the first entity in the specific cylindrical shell?
                &&
                Utilities.IsEntityInComponent((Entity)mateEntities[1]) &&                            //Is the second entity in a component?
                Utilities.IsEntityInSpecificComponent((Entity)mateEntities[1], lastCylindricalShell) &&    //Is the second entity in the last previous cylindrical shell?
                Utilities.IsEntityPlane((Entity)mateEntities[1]) &&                                         //Is the second entity a plane?
                Utilities.IsPlaneParallelToMajorPlane(solidWorksApplication, (Feature)mateEntities[1], MajorPlane.Right))   //Is the second entity parallel to assmebly's right major plane?
            {
                CurrentComponentLeftPlane = (Entity)mateEntities[0];
                PreviousComponentRightPlane = (Entity)mateEntities[1];
                return true;
            }

            //If first entity is in the last previous cylindrical shell and the second entity is in the specific cylindrical shell
            //then this mate is the mate of the cylindric shell whose left reference plane is mated with pevious cylindracl shell's right reference plane
            if (Utilities.IsEntityInComponent((Entity)mateEntities[0]) &&                            //Is the first entity in a component?
                Utilities.IsEntityInSpecificComponent((Entity)mateEntities[0], lastCylindricalShell) &&    //Is the first entity in the last previous cylindrical shell?
                Utilities.IsEntityPlane((Entity)mateEntities[0]) &&                                 //Is the first entity a plane?
                Utilities.IsPlaneParallelToMajorPlane(solidWorksApplication, (Feature)mateEntities[0], MajorPlane.Right)    ////Is the first entity parallel to assmebly's right major plane?
                &&
                (Utilities.IsEntityInComponent((Entity)mateEntities[1]) &&                          //Is the second entity in a component?
                Utilities.IsEntityInSpecificComponent((Entity)mateEntities[1], ComponentToCheck)))    //Is the second entity in the specific cylindrical shell?
            {
                PreviousComponentRightPlane = (Entity)mateEntities[0];
                CurrentComponentLeftPlane = (Entity)mateEntities[1];
                return true;
            }

            //If first entity is in a specific compoenent AND second entity is in assembly and this entity is axis
            //then this mate is the mate of the first cylindric shell whose center axis is mated with assembly's center axis
            if ((Utilities.IsEntityInComponent((Entity)mateEntities[0]) &&                          //Is the first entity in a component?
                Utilities.IsEntityInSpecificComponent((Entity)mateEntities[0], ComponentToCheck))   //Is the first entity in the specific cylindrical shell?
                &&
                (!Utilities.IsEntityInComponent((Entity)mateEntities[1]) &&                         //Is the second entity in assembly?
                Utilities.IsEntityAxis((Entity)mateEntities[1])))                                     //Is the second entity a plane?
            {
                CylindricalShellCenterAxis = (Entity)mateEntities[0];
                AssemblyCenterAxis = (Entity)mateEntities[1];
                return true;
            }

            //If first entity is in the assembly and this entity is axis AND second entity is in the specific compoenent
            //then this mate is the mate of the first cylindric shell whose center axis is mated with assembly's center axis
            if ((!Utilities.IsEntityInComponent((Entity)mateEntities[0]) &&                                     //Is the first entity in assembly?
                Utilities.IsEntityAxis((Entity)mateEntities[0])                                                  //Is the first entity axis?
                &&
                (Utilities.IsEntityInComponent((Entity)mateEntities[1]) &&                              //Is the second entity in a component?
                Utilities.IsEntityInSpecificComponent((Entity)mateEntities[1], ComponentToCheck))))      //Is the second entity in the specific cylindrical shell?
            {
                AssemblyCenterAxis = (Entity)mateEntities[0];
                CylindricalShellCenterAxis = (Entity)mateEntities[1];
                return true;
            }

            //Find specific errors

            //Error 1: Both entities are in assembly
            if (!Utilities.IsEntityInComponent((Entity)mateEntities[0])
                && !Utilities.IsEntityInComponent((Entity)mateEntities[1]))
            {
                warningService.AddError("Both entities are in the assembly.");
            }

            //Error 2: One of the entities is not in a specific component.
            if (!Utilities.IsEntityInSpecificComponent((Entity)mateEntities[0], ComponentToCheck)
               || !Utilities.IsEntityInSpecificComponent((Entity)mateEntities[1], ComponentToCheck))
            {
                warningService.AddError("One of entities is in a different component.");
            }

            return false;
        }

        /// <summary>
        ///This method copies the first cylindrical shell in the assembly and mates it with the last cylindrical shell.
        /// </summary>
        public void AddCylindricalShell()
        {
            //Get cylinder the will be copied
            Component2 cylindricalShellReference = Utilities.GetTopLevelComponents(assemblyModelDoc)[0];

            //If there is no  cylindrical shells in assembly, the copy cannot be done and method terminates
            if (cylindricalShellReference == null) 
            {
                warningService.AddError("Wasn't able to get first cylindrical shell. Check if it exists.");
                return;
            }

            //Create a copy cylinder shell of reference cylinder shell.
            cylindricalShell = assemblyDocument.AddComponent5(cylindricalShellReference.GetPathName(), 0, "", false, "", 0, 0, 0);

            //If the component is not added, the method is terminated
            if (cylindricalShell is null)
            {
                warningService.AddError("Wans't able to copy the first cylindrical shell.");
                return;
            }

            //Select cylinder shell copy to make it independent
            Feature cylindricalShellFeature = Utilities.GetNthFeatureOfType(assemblyModelDoc, FeatureType.Component, Utilities.GetTopLevelComponents(assemblyModelDoc).Count);
            Utilities.SelectFeature(assemblyModelDoc, cylindricalShellFeature);

            //Make the cylinder shell copy independent
            assemblyDocument.MakeIndependent(CYLINDRICAL_SHELL_PATH);

            //Rename cylindrical shell copy
            Utilities.RenameFeature(cylindricalShellFeature, CYLINDRICAL_SHELL_NAME + " " + Utilities.GetTopLevelComponents(assemblyModelDoc).Count);

            //Assign right and left reference planes of cylidrical shell copy
            rightReferencePlane = Utilities.GetNthFeatureOfType(cylindricalShell, FeatureType.RefPlane, 5);
            leftReferencePlane = Utilities.GetNthFeatureOfType(cylindricalShell, FeatureType.RefPlane, 4);

            //Mate cylindrical shell

            //Get list of cylindrical shells in the assemlby.
            List<Component2> listOfCylindricalShells = Utilities.GetTopLevelComponents(assemblyModelDoc);

            //Get last previous cylindrical shell's right plane
            Component2 lastPreviousCylindricalShell = listOfCylindricalShells[(listOfCylindricalShells.Count - 2)];
            Feature previousCylindricalShellRightPlane = Utilities.GetNthFeatureOfType(lastPreviousCylindricalShell, FeatureType.RefPlane, 5);

            //Mates new cylindrical shell's left reference plane with the last one's right reference plane.
            mateCoincidentLeftPlane.CreateMate(
                ComponentFeature1: previousCylindricalShellRightPlane,
                ComponentFeature2: leftReferencePlane,
                AlignmentType: MateAlignment.Anti_Aligned,
                Name: CYLINDRICAL_SHELL_NAME + listOfCylindricalShells.Count + " - " + LEFT_REFERENCE_PLANE_NAME);

            //Mates new cylindrical shell's center axis with assembly's center axis
            mateCoincidentAxis.CreateMate(
                ComponentFeature1: Utilities.GetNthFeatureOfType(assemblyModelDoc, FeatureType.RefAxis, 1),
                ComponentFeature2: Utilities.GetNthFeatureOfType(cylindricalShell, FeatureType.RefAxis, 1),
                AlignmentType: MateAlignment.Anti_Aligned,
                Name: CYLINDRICAL_SHELL_NAME + listOfCylindricalShells.Count + " - " + CENTER_AXIS_NAME);

            //Mates new cylindrical shell's front plane with assembly's front plane with angle.
            mateAngleFrontPlane.CreateMate(
                ExternalEntity: (Entity)Utilities.GetMajorPlane(assemblyModelDoc, MajorPlane.Front),
                ComponentEntity: (Entity)Utilities.GetMajorPlane(cylindricalShell, MajorPlane.Front),
                ReferenceEntity: (Entity)Utilities.GetNthFeatureOfType(assemblyModelDoc, FeatureType.RefAxis, 1),
                Angle: 0.78539816339744830961566084581988,
                FlipDimension: false,
                Name: CYLINDRICAL_SHELL_NAME + listOfCylindricalShells.Count + " - " + FRONT_REFERENCE_PLANE_NAME);
        }

        /// <summary>
        /// Delete the cylindrical shell component
        /// </summary>
        public void Remove()
        {
            SelectionMgr selectionManager = (SelectionMgr)assemblyModelDoc.SelectionManager;
            SelectData selectData = selectionManager.CreateSelectData();

            //Select the dished end to be deleted
            cylindricalShell.Select4(false, selectData, false);

            //Delete selected dished end
            assemblyDocument.DeleteSelections(0);
        }

        public void ChangeLength(double Length)
        {
            ModelDocExtension modelDocExtension = assemblyModelDoc.Extension;

            modelDocExtension.SelectByID2(
                "D3@Sketch1@" + cylindricalShell.Name2 + "@" + assemblyModelDoc.GetTitle(),
                "DIMENSION",
                0, 0, 0, false, 1, null, 0);

            Dimension myDimension = (Dimension)assemblyModelDoc.Parameter("D3@Sketch1@" + cylindricalShell.Name2 + ".Part");

            myDimension.SetSystemValue3(2, 1, "");

            assemblyDocument.ForceRebuild2(true);

            assemblyModelDoc.ClearSelection2(true);
        }
    }

    
}
