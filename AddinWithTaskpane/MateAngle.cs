using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

namespace AddinWithTaskpane
{
    internal class MateAngle
    {
        private ModelDoc2 assemblyModelDoc;
        private AssemblyDoc assemblyDoc;

        private Feature mate;
        private MateFeatureData mateFeatureData;
        private AngleMateFeatureData angleMateFeatureData;

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
        /// <param name="AssemblyDocument"></param>
        public MateAngle(ModelDoc2 AssemblyDocument)
        {
            assemblyModelDoc = AssemblyDocument;
            assemblyDoc = (AssemblyDoc)assemblyModelDoc;

            mateFeatureData = assemblyDoc.CreateMateData((int)swMateType_e.swMateANGLE);
            angleMateFeatureData = (AngleMateFeatureData)mateFeatureData;
        }

        public MateAngle(ModelDoc2 AssemblyDocument, Feature InputMate) // sito konstruktoriaus nebandziau.
        {
            assemblyModelDoc = AssemblyDocument;
            mate = InputMate;
            mateFeatureData = (MateFeatureData)mate.GetDefinition();
            angleMateFeatureData = (AngleMateFeatureData)mateFeatureData;
        }


        /// <summary>
        /// Creates mate
        /// </summary>
        /// <param name="ModelDocument"></param>
        /// <param name="AssemblyFeature"></param>
        /// <param name="Component"></param>
        /// <param name="ComponentEntity"></param>
        public void CreateMate(
            Entity ExternalEntity,
            Entity ComponentEntity,
            Entity ReferenceEntity,
            double Angle,
            bool FlipDimension,
            string Name)
        {
            angleMateFeatureData.FlipDimension = FlipDimension;
            angleMateFeatureData.Angle = Angle;
            angleMateFeatureData.EntitiesToMate = new Entity[] { ExternalEntity, ComponentEntity };
            angleMateFeatureData.ReferenceEntity = ReferenceEntity;

            //Get last components to get its angle mate's dimension
            Component2 lastCylinderShell = Utilities.GetTopLevelComponents(assemblyModelDoc)[Utilities.GetTopLevelComponents(assemblyModelDoc).Count - 2];
            string name = lastCylinderShell.Name;

            mate = assemblyDoc.CreateMate(angleMateFeatureData);
            mate.Name = Name;
        }
        public void FlipDimension()
        {
            angleMateFeatureData.FlipDimension = !angleMateFeatureData.FlipDimension;
            mate.ModifyDefinition(angleMateFeatureData, assemblyDoc, null);
        }
    }
}
