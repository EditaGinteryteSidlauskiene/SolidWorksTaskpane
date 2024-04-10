using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System.Diagnostics;

namespace AddinWithTaskpane
{
    internal class MateDistance
    {
        private ModelDoc2 assemblyModelDoc;
        private AssemblyDoc assemblyDoc;
        private ModelDocExtension modelDocExtension;

        private Feature mate;
        private SelectData selectData;

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
        public MateDistance(ModelDoc2 inputswDoc) => InitialAssignment(inputswDoc);

        public MateDistance(ModelDoc2 inputswDoc, Feature InputMate) : this(inputswDoc) => mate = InputMate;

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
        /// Edit mate and clear selection
        /// </summary>
        /// <param name="AlignmentType"></param>
        /// <param name="Distance"></param>
        private void EditMate()
        {
            //Select mate
            mate.Select2(true, 1);

            string mateName = mate.Name;

            // Cast the MateFeatureData object to a CoincidentMateFeatureData object. 
            DistanceMateFeatureData distMateData = (DistanceMateFeatureData)mate.GetDefinition();

            //Edit mate
            ((AssemblyDoc)assemblyModelDoc).EditMate4(
                MateTypeFromEnum: 5,
                AlignFromEnum: distMateData.MateAlignment,
                Flip: false,
                Distance: distMateData.Distance,
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
                ErrorStatus: out _);

            //Assign editted mate
            mate = Utilities.GetMateByName(assemblyModelDoc, mateName);

            //Clear selection
            assemblyModelDoc.ClearSelection2(true);
        }

        public void CreateMate(Feature ComponentFeature1, Feature ComponentFeature2, MateAlignment AlignmentType, double Distance, string Name) =>
            CreateMate((Entity)ComponentFeature1, (Entity)ComponentFeature2, AlignmentType, Distance, Name);

        /// <summary>
        /// Creates a new mate.
        /// </summary>
        /// <param name="ComponentFeature1"></param>
        /// <param name="ComponentFeature2"></param>
        /// <param name="AlignmentType"></param>
        /// <param name="Distance"></param>
        /// <param name="Name"></param>
        public void CreateMate(Entity ComponentFeature1, Entity ComponentFeature2,MateAlignment AlignmentType, double Distance, string Name)
        {
            //Select entities
            ComponentFeature1.Select4(false, selectData);
            ComponentFeature2.Select4(true, selectData);

            //Create Mate
            mate = (Feature)assemblyDoc.AddDistanceMate(
                AlignFromEnum: (int)AlignmentType,
                Flip: false,
                Distance: Distance,
                DistanceAbsUpperLimit: Distance,
                DistanceAbsLowerLimit: Distance,
                FirstArcCondition: 0,
                SecondArcCondition: 0,
                ErrorStatus: out _);

            //Assign name for the mate
            mate.Name = Name;

            assemblyModelDoc.ClearSelection2(true);
        }

        /// <summary>
        /// Change alignment of a mate
        /// </summary>
        /// <param name="Mate"></param>
        public void ChangeAlignment()
        {
            // Access mate's feature and get the MateFeatureData object. 
            IMateFeatureData mateData = mate.GetDefinition();

            // Cast the MateFeatureData object to a CoincidentMateFeatureData object. 
            IDistanceMateFeatureData distMateData = (IDistanceMateFeatureData)mateData;

            // Change the alignment. 0 - ALIGNED
            distMateData.MateAlignment = (distMateData.MateAlignment == 0 ? (int)swMateAlign_e.swMateAlignANTI_ALIGNED : (int)swMateAlign_e.swMateAlignALIGNED);

            // Updates the definition of a feature with the new values in an associated feature data object 
            mate.ModifyDefinition(mateData, assemblyModelDoc, null);
        }

        /// <summary>
        /// Changes distance property of the mate
        /// </summary>
        /// <param name="NewDistance"></param>
        public void ChangeDistance(double NewDistance)
        {
            // Access mate's feature and get the MateFeatureData object. 
            IMateFeatureData mateData = mate.GetDefinition();

            // Cast the MateFeatureData object to a CoincidentMateFeatureData object. 
            IDistanceMateFeatureData mateToEdit = (IDistanceMateFeatureData)mateData;

            mateToEdit.Distance = NewDistance;

            // Updates the definition of a feature with the new values in an associated feature data object 
            mate.ModifyDefinition(mateData, assemblyModelDoc, null);
        }

        /// <summary>
        /// Moves entities to opposite sides of the dimension of this distance mate
        /// </summary>
        public void FlipDimension()
        {
            // Access mate's feature and get the MateFeatureData object. 
            IMateFeatureData mateData = mate.GetDefinition();

            // Cast the MateFeatureData object to a CoincidentMateFeatureData object. 
            IDistanceMateFeatureData mateToEdit = (IDistanceMateFeatureData)mateData;

            mateToEdit.FlipDimension = true;

            // Updates the definition of a feature with the new values in an associated feature data object 
            mate.ModifyDefinition(mateData, assemblyModelDoc, null);

        }

        /// <summary>
        /// Changes one entity in the mate, when 2 entities are given
        /// </summary>
        /// <param name="NewAssemblyFeature"></param>
        /// <param name="ComponentInMate"></param>
        /// <param name="ComponentFeature"></param>
        /// <param name="Distance"></param>
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
        /// <param name="Component1"></param>
        /// <param name="ComponentFeature1"></param>
        /// <param name="Component2"></param>
        /// <param name="ComponentFeature2"></param>
        /// <param name="Distance"></param>
        public void ChangeMateEntity(Feature Feature1, Feature Feature2)
        {
            // Select features
            Feature1.Select2(false, 1);
            Feature2.Select2(true, 1);

            //Edit mate
            EditMate();
        }

        /// <summary>
        /// Changes one entity in the mate, when 1 feature and 1 entity are given
        /// </summary>
        /// <param name="Feature"></param>
        /// <param name="Entity"></param>
        public void ChangeMateEntity(Feature Feature, Entity Entity)
        {
            // Select features
            //Select entities and mate
            Feature.Select2(false, 1);
            Entity.Select4(true, selectData);

            //Edit mate
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
