using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System.Diagnostics;

namespace AddinWithTaskpane
{
    internal class MateCoincident
    {
        private ModelDoc2 assemblyModelDoc;
        private AssemblyDoc assemblyDoc;
        private ModelDocExtension modelDocExtension;

        private Feature mate;
        SelectData selectData;

        /// <summary>
        /// Check if the mate exists
        /// </summary>
        /// <returns></returns>
        public bool IsSet()
        {
            return (assemblyModelDoc != null) && (mate != null);
        }

        /// <summary>
        /// Constructor to create a new mate object
        /// </summary>
        /// <param name="inputswDoc"></param>
        public MateCoincident(ModelDoc2 inputswDoc) => InitialAssignment(inputswDoc);

        public MateCoincident(ModelDoc2 inputswDoc, Feature InputMate) : this(inputswDoc) => mate = InputMate;

        /// <summary>
        /// Assign document, assembly, selection manager
        /// </summary>
        /// <param name="inputswDoc"></param>
        private void InitialAssignment(ModelDoc2 inputswDoc)
        {
            assemblyModelDoc = inputswDoc;
            assemblyDoc = (AssemblyDoc)assemblyModelDoc;
            modelDocExtension = assemblyModelDoc.Extension;

            //Allows to select objects.
            SelectionMgr selectionManager = (SelectionMgr)assemblyModelDoc.SelectionManager;
            selectData = selectionManager.CreateSelectData();
            selectData.Mark = 1;
        }

        /// <summary>
        /// Suppress mate
        /// </summary>
        public void Suppress()
        {
            Utilities.ChangeSuppression(mate, 0);
        }

        /// <summary>
        /// Unsuppress mate
        /// </summary>
        public void Unsuppress()
        {
            Utilities.ChangeSuppression(mate, 1);
        }

        /// <summary>
        /// Edits mate
        /// </summary>
        private void EditMate()
        {
            mate.Select2(true, 1);

            string mateName = mate.Name;

            // Cast the MateFeatureData object to a CoincidentMateFeatureData object. 
            CoincidentMateFeatureData coincMateData = (CoincidentMateFeatureData)mate.GetDefinition();

            //Edit mate
            assemblyDoc.EditMate4(
                MateTypeFromEnum: 0,
                AlignFromEnum: coincMateData.MateAlignment,
                Flip: false,
                Distance: 0,
                DistanceAbsUpperLimit: 0,
                DistanceAbsLowerLimit: 0,
                GearRatioNumerator: 0,
                GearRatioDenominator: 0,
                Angle: 0,
                AngleAbsUpperLimit: 0,
                AngleAbsLowerLimit: 0,
                ForPositioningOnly: false,
                LockRotation: true,
                WidthMateOption: 0,
                RepairMatesWithSameMissingEntity: false,
                ErrorStatus: out int _);

            //Assign editted mate
            mate = Utilities.GetMateByName(assemblyModelDoc, mateName);

            //Clear selection
            assemblyModelDoc.ClearSelection2(true);
        }

        /// <summary>
        /// Creates mate
        /// </summary>
        /// <param name="ComponentFeature1"></param>
        /// <param name="ComponentFeature2"></param>
        /// <param name="AlignmentType"></param>
        /// <param name="Name"></param>
        public void CreateMate(Feature ComponentFeature1, Feature ComponentFeature2, MateAlignment AlignmentType, string Name) =>
           CreateMate((Entity)ComponentFeature1, (Entity)ComponentFeature2, AlignmentType, Name);

        /// <summary>
        /// Creates mate
        /// </summary>
        /// <param name="ModelDocument"></param>
        /// <param name="AssemblyFeature"></param>
        /// <param name="Component"></param>
        /// <param name="ComponentFeature"></param>
        public void CreateMate(Entity ComponentFeature1, Entity ComponentFeature2, MateAlignment AlignmentType, string Name)
        {
            //Creates a mate feature data object for the specified mate type. This is required to access CreateMate method
            CoincidentMateFeatureData coincidentMateFeatureData = (CoincidentMateFeatureData)assemblyDoc.CreateMateData((int)swMateType_e.swMateCOINCIDENT);

            //NEVEIKIA --------------------------------------------------------------------------------------
            //coincidentMateFeatureData.EntitiesToMate = new Entity[] { ComponentFeature1, ComponentFeature1 };
            //NEVEIKIA --------------------------------------------------------------------------------------

            // VEIKIA ---------------------------------------------------------------------------------------
            //Select entities
            ComponentFeature1.Select4(false, selectData);
            ComponentFeature2.Select4(true, selectData);
            // VEIKIA ---------------------------------------------------------------------------------------

            //Alignment - Aligned
            coincidentMateFeatureData.MateAlignment = (int)AlignmentType;

            //Create Mate
            mate = (Feature)assemblyDoc.CreateMate(coincidentMateFeatureData);

            //Assign name for the mate
            mate.Name = Name;

            //Clear selection
            assemblyModelDoc.ClearSelection2(true);
        }

        /// <summary>
        /// Change alignment of a mate
        /// </summary>
        /// <param name="MateToChangeAlignment"></param>
        public bool ChangeAlignment()
        {
            // Access mate's feature and get the MateFeatureData object. 
            IMateFeatureData mateData = mate.GetDefinition();

            // Cast the MateFeatureData object to a CoincidentMateFeatureData object. 
            CoincidentMateFeatureData coincMateData = (CoincidentMateFeatureData)mateData;

            // Change the alignment. 0 - ALIGNED
            coincMateData.MateAlignment = (coincMateData.MateAlignment == 0 ? (int)swMateAlign_e.swMateAlignANTI_ALIGNED : (int)swMateAlign_e.swMateAlignALIGNED);

            // Updates the definition of a feature with the new values in an associated feature data object 
            return mate.ModifyDefinition(mateData, assemblyModelDoc, null);
        }

        /// <summary>
        /// Changes one entity in the mate, when 2 entities are given
        /// </summary>
        /// <param name="Entity1"></param>
        /// <param name="Entity2"></param>
        public void ChangeMateEntity(Entity Entity1, Entity Entity2)
        {
            //Select entities and mate
            Entity1.Select4(false, selectData);
            Entity2.Select4(true, selectData);

            EditMate();
        }

        /// <summary>
        /// Changes one entity in the mate, when 2 features are given
        /// </summary>
        /// <param name="Entity1"></param>
        /// <param name="Entity2"></param>
        public void ChangeMateEntity(Feature Feature1, Feature Feature2)
        {
            Feature1.Select2(false, 1);
            Feature2.Select2(true, 1);

            EditMate();
        }

        /// <summary>
        /// Changes one entity in the mate, when 1 feature and 1 entity are given
        /// </summary>
        /// <param name="Entity1"></param>
        /// <param name="Entity2"></param>
        public void ChangeMateEntity(Feature Feature, Entity Entity)
        {
            //Select entities and mate
            Feature.Select2(false, 1);
            Entity.Select4(true, selectData);

            EditMate();
        }

        /// <summary>
        /// Delete selected mate
        /// </summary>
        public void Delete()
        {
            //Select mate
            mate.Select2(false, 1);

            //Delete
            modelDocExtension.DeleteSelection2(0);
        }
    }
}
