using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;

namespace AddinWithTaskpane
{

    internal class Utilities
    {
        /// <summary>
        /// Dictionary of Feature Type Names. It can be used to get Feature Type Name as a string
        /// </summary>
        /// 
        static readonly Dictionary<FeatureType, string> FeatureNames = new Dictionary<FeatureType, string> ()
        {
            {FeatureType.RefPlane, "RefPlane" },
            {FeatureType.RefAxis, "RefAxis" },
            {FeatureType.Component, "Reference" },
        };


        /// <summary>
        /// Dictionary of Entity Types. These type names are used in SelectByID2 method. The key in 
        /// this dictionary is feature's type name.
        /// </summary>
        public static Dictionary<string, string> EntityType = new Dictionary<string, string>()
        {
            {"RefPlane", "PLANE" },
            {"RefAxis", "AXIS" },
            {"MateCoincident", "MATE" },
            {"MateDistanceDim", "MATE"},
            {"Reference", "COMPONENT" }
        };


        /// <summary>
        /// Gets feature by name by iterating all features until the requested one is reached.
        /// </summary>
        /// <param name="ModelDocument"></param>
        /// <param name="Name"></param>
        /// <returns></returns>
        public static Feature GetFeatureByName(ModelDoc2 ModelDocument, string Name)
        {
            //Starting from the first feature
            Feature loopFeature = ModelDocument.IFirstFeature();

            //Loop features until the requested feature is found
            while (loopFeature != null)
            {
                if (loopFeature.Name == Name)
                {
                    return loopFeature;
                }

                //Get next feature
                loopFeature = (Feature)loopFeature.GetNextFeature();
            }

            return null;
        }

        public static Feature GetFeatureByName(ModelDoc2 ModelDocument, Component2 Component, string Name)
        {
            //Starting from the first feature
            Feature loopFeature = Component.FirstFeature();

            //Loop features until the requested feature is found
            while (loopFeature != null)
            {
                if (loopFeature.Name == Name)
                {
                    return loopFeature;
                }

                //Get next feature
                loopFeature = (Feature)loopFeature.GetNextFeature();
            }

            return null;
        }

        /// <summary>
        /// Renames feature of the assembly
        /// </summary>
        /// <param name="swFeature"></param>
        /// <param name="Name"></param>
        public static void RenameFeature(Feature swFeature, string Name) => swFeature.Name = Name;

        /// <summary>
        /// Returns a major plane of assembly that is requested by providing its type from MajorPlane enum. Loops all features
        /// until the requested one is found.
        /// </summary>
        /// <param name="PlaneType"></param>
        /// <param name="ModelDocument"></param>
        /// <returns></returns>
        public static Feature GetMajorPlane(ModelDoc2 ModelDocument, MajorPlane PlaneType) => GetNthFeatureOfType(ModelDocument.IFirstFeature(), FeatureType.RefPlane, (int)PlaneType);

        /// <summary>
        /// Returns a componnent's major plane that is requested by providing its type from MajorPlane enum and component. 
        /// Loops all features until the requested one is found.
        /// </summary>
        /// <param name="PlaneType"></param>
        /// <param name="Component"></param>
        /// <returns></returns>
        public static Feature GetMajorPlane(Component2 Component, MajorPlane PlaneType) => GetNthFeatureOfType(Component.FirstFeature(), FeatureType.RefPlane, (int)PlaneType);

        public static Feature GetMateByName(ModelDoc2 ModelDocument, string MateName)
        {
            Feature loopFeature = ModelDocument.IFirstFeature();

            //Loop features until the requested feature is found
            while (loopFeature != null)
            {
                if (loopFeature.GetTypeName2() == "MateGroup")
                {
                    Feature subFeature = (Feature)loopFeature.GetFirstSubFeature();

                    while (subFeature != null)
                    {
                        if (subFeature.Name == MateName)
                        {
                            Feature mate = subFeature;
                            return mate;
                        }

                        subFeature = (Feature)subFeature.GetNextSubFeature();
                    }
                    return null;
                }

                //Get next feature
                loopFeature = (Feature)loopFeature.GetNextFeature();
            }
            return null;
        }

        /// <summary>
        /// Creates and renames an assembly reference plane that is in a specified
        /// distance of the selected major plane.
        /// </summary>
        /// <param name="ModelDocument"></param>
        /// <param name="ExistingPlane"></param>
        /// <param name="Distance"></param>
        /// <param name="Name"></param>
        /// <returns></returns>
        public static Feature CreateReferencePlaneWithDistance(
            ModelDoc2 ModelDocument,
            Feature ExistingPlane,
            double Distance,
            string Name)
        {
            //Selects the existing plane 
            _ = ExistingPlane.Select2(false, 0);

            //Creates a new reference plane
            Feature ReferencePlane = (Feature)ModelDocument.FeatureManager.InsertRefPlane(8, Distance, 0, 0, 0, 0);

            //Rename just created reference plane
            RenameFeature(ReferencePlane, Name);

            return ReferencePlane;
        }

        /// <summary>
        /// Changes the distance of reference plane from the starting plane
        /// </summary>
        /// <param name="ModelDocument"></param>
        /// <param name="ReferencePlane"></param>
        /// <param name="Distance"></param>
        public static bool ChangeDistanceOfReferencePlane(
            ModelDoc2 ModelDocument,
            Feature ReferencePlane,
            double Distance)
        {
            //Get access to reference plane properties
            RefPlaneFeatureData referencePlaneFeatureData;
            referencePlaneFeatureData = ReferencePlane.GetDefinition();

            //Set new distance
            referencePlaneFeatureData.Distance = Distance;

            //Modify changes
            return ReferencePlane.ModifyDefinition(referencePlaneFeatureData, ModelDocument, null);
        }

        /// <summary>
        /// Changes reference entity (NewReferencePlane) for reference plane feature (ReferencePlane).
        /// </summary>
        /// <param name="ModelDocument"></param>
        /// <param name="NewReferencePlane"></param>
        /// <param name="ReferencedPlane"></param>
        public static bool ChangeReferenceOfReferencePlane(
            ModelDoc2 ModelDocument,
            Feature NewReferencePlane,
            Feature ReferencedPlane)
        {
            //Get access to reference plane properties
            RefPlaneFeatureData referencePlaneFeatureData;
            referencePlaneFeatureData = ReferencedPlane.GetDefinition();

            //Set new reference plane
            referencePlaneFeatureData.Reference[0] = NewReferencePlane;

            //Modify changes
            return ReferencedPlane.ModifyDefinition(referencePlaneFeatureData, ModelDocument, null);
        }

        /// <summary>
        /// Gets last feature of a certain type in Feature Manager.
        /// </summary>
        /// <param name="ModelDocument"></param>
        /// <param name="DesiredFeatureType"></param>
        /// <returns></returns>
        public static Feature GetLastFeatureOfType(Feature FirstFeature, FeatureType DesiredFeatureType)
        {
            // Getting the name of feature type as a string
            string featureTypeName = FeatureNames[DesiredFeatureType];

            // Getting the first feature
            Feature loopFeature = FirstFeature;

            //Last feature of a type has not been found yet
            Feature lastFeatureOfType = null;

            //Loops features
            while (loopFeature != null)
            {
                //Checking if the current feature is of the same type as the requested one.
                if (featureTypeName == loopFeature.GetTypeName2())
                {
                    //Setting the current feature as the last one.
                    lastFeatureOfType = loopFeature;
                }
                //Go to the next feature
                loopFeature = (Feature)loopFeature.GetNextFeature();
            }
            return lastFeatureOfType;
        }

        public static Feature GetLastFeatureOfType(ModelDoc2 ModelDocument, FeatureType DesiredFeatureType) => GetLastFeatureOfType(ModelDocument.IFirstFeature(), DesiredFeatureType);

        public static Feature GetLastFeatureOfType(Component2 Component, FeatureType DesiredFeatureType) => GetLastFeatureOfType(Component.FirstFeature(), DesiredFeatureType);

        private static Feature GetNthFeatureOfType(Feature FirstFeature, FeatureType DesiredFeatureType, int Count)
        {
            // Error Handling: Ensure the provided count is valid
            if (Count <= 0)
            {
                return null;
            }

            // Getting the desired feature type's name for comparison
            string featureTypeName = FeatureNames[DesiredFeatureType];

            // Initialize variables 
            Feature loopFeature = FirstFeature;
            int featureCounter = 0;

            // Iterate through features 
            while (loopFeature != null)
            {
                if (featureTypeName == loopFeature.GetTypeName2())
                {
                    featureCounter++;

                    if (featureCounter == Count)
                        return loopFeature;
                }

                loopFeature = (Feature)loopFeature.GetNextFeature();
            }

            return null;
        }

        /// <summary>
        /// Retrieves the 'nth' feature of a specific type within a ModelDoc2 document.
        /// </summary>
        /// <param name="ModelDocument">The model document containing the features.</param>
        /// <param name="DesiredFeatureType">The desired feature type to search for.</param>
        /// <param name="Count">The ordinal position (1-based) of the desired feature.</param>
        /// <returns>The 'nth' Feature object of the specified type, or null if not found or the count is invalid.</returns>
        public static Feature GetNthFeatureOfType(ModelDoc2 ModelDocument, FeatureType DesiredFeatureType, int Count) => GetNthFeatureOfType(ModelDocument.IFirstFeature(), DesiredFeatureType, Count);

        /// <summary>
        /// Retrieves the 'nth' feature of a specific type within a ModelDoc2 document.
        /// </summary>
        /// <param name="Component">The component document containing the features.</param>
        /// <param name="DesiredFeatureType">The desired feature type to search for.</param>
        /// <param name="Count">The ordinal position (1-based) of the desired feature.</param>
        /// <returns>The 'nth' Feature object of the specified type, or null if not found or the count is invalid.</returns>
        public static Feature GetNthFeatureOfType(Component2 Component, FeatureType DesiredFeatureType, int Count) => GetNthFeatureOfType(Component.FirstFeature(), DesiredFeatureType, Count);

        public static void SelectFeature(ModelDoc2 ModelDocument, Feature Feature)
        {
            ModelDocExtension swModelDocExt = ModelDocument.Extension;

            //Setting feature type name
            string featureTypeName = EntityType[Feature.GetTypeName2()];

            // Select assembly's feature
            swModelDocExt.SelectByID2(
                Name: Feature.Name,
                Type: featureTypeName,
                X: 0,
                Y: 0,
                Z: 0,
                Append: true,
                Mark: 1,
                Callout: null,
                SelectOption: (int)swSelectOption_e.swSelectOptionDefault);
        }

        public static void SelectFeature(ModelDoc2 ModelDocument, Component2 Component, Feature ComponentFeature)
        {
            ModelDocExtension swModelDocExt = ModelDocument.Extension;

            //Getting full entity name, which will be used later to select the entity. Example of full entity name
            //FullEntityName = "Top@" + strComponentName + "@" + AssemblyName;
            string fullComponentFeatureName =
                ComponentFeature.Name
                + "@"
                + Component.Name
                + "@"
                + ModelDocument.GetTitle();

            //Setting entity type name
            string componentFeatureTypeName = EntityType[ComponentFeature.GetTypeName2()];
           
            // Select component's feature
            swModelDocExt.SelectByID2(
                Name: fullComponentFeatureName,
                Type: componentFeatureTypeName,
                X: 0,
                Y: 0,
                Z: 0,
                Append: true,
                Mark: 1,
                Callout: null,
                SelectOption: (int)swSelectOption_e.swSelectOptionDefault);
        }

        public static void InsertCenterAxis(ModelDoc2 ModelDocument, MajorPlane Plane1, MajorPlane Plane2, string Name)
        {
            GetMajorPlane(ModelDocument, Plane1).Select2(false, 0);
            GetMajorPlane(ModelDocument, Plane2).Select2(true, 0);

            ModelDocument.InsertAxis2(true);

            Feature CenterAxis = GetLastFeatureOfType(ModelDocument, FeatureType.RefAxis);
            RenameFeature(CenterAxis, Name);

            ModelDocument.ClearSelection2(true);
        }

        public static Component2 AddComponent(ISldWorks SolidWorksApplication, ModelDoc2 AssemblyDocument, string ComponentPath)
        {
            // Open the document of the component to be added
            SolidWorksApplication.OpenDoc6(ComponentPath, (int)swDocumentTypes_e.swDocPART, (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", 1, 1);

            // Add the part to the assembly document
            Component2 Component = ((AssemblyDoc)AssemblyDocument).AddComponent5(ComponentPath, (int)swAddComponentConfigOptions_e.swAddComponentConfigOptions_CurrentSelectedConfig, "", false, "", 1, 0, 0);

            // Close the document of added component
            SolidWorksApplication.CloseDoc(ComponentPath);

            // Making the added component float
            // Checking if the component is fixed
            if (Component.IsFixed() == true)
            {
                // Selecting the component
                Component.Select2(false, 0);

                // Unfixing component
                ((AssemblyDoc)AssemblyDocument).UnfixComponent();
            }

            return Component;
        }

        /// <summary>
        /// Change the state of suppression. Suppression State = 0 to suppress, 1 to unsuppress.
        /// </summary>
        /// <param name="Mate"></param>
        /// <param name="SuppressionState"></param>
        public static void ChangeSuppression(Feature Feature, int SuppressionState)
        {
            Feature.SetSuppression2(SuppressionState, 2, "");
        }

        public static void SupressFeature(Feature Feature) => ChangeSuppression(Feature, 0);

        public static void UnsupressFeature(Feature Feature) => ChangeSuppression(Feature, 1);


        public static void SupressComponent(Component2 Component) => Component.SetSuppression((int)swComponentSuppressionState_e.swComponentSuppressed);

        public static void UnsupressComponent(Component2 Component) => Component.SetSuppression((int)swComponentSuppressionState_e.swComponentFullyResolved);


        /// <summary>
        /// Adds a coincident mate between two features in a SOLIDWORKS assembly, and assigns a descriptive name to the mate.
        /// </summary>
        /// <param name="Assembly">The SOLIDWORKS assembly document in which to create the mate.</param>
        /// <param name="Component1">The first component involved in the mate.</param>
        /// <param name="Component1Feature">The feature of the first component to be mated.</param>
        /// <param name="Component2">The second component involved in the mate.</param>
        /// <param name="Component2Feature">The feature of the second component to be mated.</param>
        public static void AddMateCoincident(ModelDoc2 Assembly, Component2 Component1, Feature Component1Feature, Component2 Component2, Feature Component2Feature)
        {
            // Select the features for mating
            SelectFeature(Assembly, Component1, Component1Feature);
            SelectFeature(Assembly, Component2, Component2Feature);

            // Create the coincident mate
            Feature mate = (Feature)((AssemblyDoc)Assembly).AddMate5(
                MateTypeFromEnum: (int)swMateType_e.swMateCOINCIDENT,
                AlignFromEnum: (int)swMateAlign_e.swMateAlignALIGNED,
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
                LockRotation: false,
                WidthMateOption: 0,
                ErrorStatus: out _);

            // Generate a descriptive name for the mate
            string mateName = "Coincident " +
                              Component1Feature.Name + "_" + Component1.Name2 + "_" +
                              Component2Feature.Name + "_" + Component2.Name2;

            // Assign the name to the mate
            RenameFeature(mate, mateName);
        }


        /// <summary>
        /// Looks for a custom text property, if there is one, changes its value. If the custom text property
        /// does not exist, creates a new one.
        /// </summary>
        /// <param name="SolidWorksApplication"></param>
        /// <param name="ModelDocument"></param>
        public static void AddOrEditCustomTextProperty(ModelDoc2 ModelDocument, string CustomPropertyName, string NewCustomPropertyValue)
        {
            //Get the custom property manager
            CustomPropertyManager customPropertyManager = ModelDocument.Extension.CustomPropertyManager[""];

            // Store the property type
            int propertyType = customPropertyManager.GetType2(CustomPropertyName);

            /* Gets the value and the evaluated value of the specified custom property.
             Result codes when getting custom properties: 0 = Cached value was returned,
            1 = Custom property does not exist, 2 = Resolved value was returned.*/
            int status = customPropertyManager.Get6(
                FieldName: CustomPropertyName, 
                UseCached: false,
                ValOut: out _,
                ResolvedValOut: out _,
                WasResolved: out _,
                LinkToProperty: out _);

            // If the custom property exists
            if (status == 2 && propertyType == 30)
            {
                // Change the value of the custom property
                customPropertyManager.Add3(
                    FieldName: CustomPropertyName,
                    FieldType: 30,
                    FieldValue: NewCustomPropertyValue,
                    OverwriteExisting: (int)swCustomPropertyAddOption_e.swCustomPropertyReplaceValue);
            }
            // If the custom property exists, but its type is not 'text'
            else if (status == 2)
            {
                // Add a new custom property of a type 'text' and delete previously found one.
                customPropertyManager.Add3(
                    FieldName: CustomPropertyName,
                    FieldType: 30,
                    FieldValue: NewCustomPropertyValue,
                    OverwriteExisting: (int)swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd);
            }
            else
            {
                // Creates a new custom property
                customPropertyManager.Add3(
                    FieldName: CustomPropertyName,
                    FieldType: 30,
                    FieldValue: NewCustomPropertyValue,
                    OverwriteExisting: (int)swCustomPropertyAddOption_e.swCustomPropertyOnlyIfNew);
            }
        }


        /// <summary>
        /// Looks for a custom number property, if there is one, changes its value. If the custom number property
        /// does not exist, creates a new one.
        /// </summary>
        /// <param name="SolidWorksApplication"></param>
        /// <param name="ModelDocument"></param>
        public static void AddOrEditCustomNumberProperty(ModelDoc2 ModelDocument, string CustomPropertyName, double NewCustomPropertyValue)
        {
            //Get the custom property manager
            CustomPropertyManager customPropertyManager = ModelDocument.Extension.CustomPropertyManager[""];

            // Store the property type
            int propertyType = customPropertyManager.GetType2(CustomPropertyName);

            /* Gets the value and the evaluated value of the specified custom property.
             Result codes when getting custom properties: 0 = Cached value was returned,
            1 = Custom property does not exist, 2 = Resolved value was returned.*/
            int status = customPropertyManager.Get6(
                FieldName: CustomPropertyName,
                UseCached: false,
                ValOut: out string _,
                ResolvedValOut: out string _,
                WasResolved: out _,
                LinkToProperty: out _);

            // If the custom property exists
            if (status == 2 && propertyType == 3)
            {
                // Change the value of the custom property
                customPropertyManager.Add3(
                    FieldName: CustomPropertyName,
                    FieldType: 3,
                    FieldValue: NewCustomPropertyValue.ToString(),
                    OverwriteExisting: (int)swCustomPropertyAddOption_e.swCustomPropertyReplaceValue);
            }
            // If the custom property exists, but its type is not 'text'
            else if (status == 2)
            {
                // Add a new custom property of a type 'text' and delete previously found one.
                customPropertyManager.Add3(
                    FieldName: CustomPropertyName,
                    FieldType: 3,
                    FieldValue: NewCustomPropertyValue.ToString(),
                    OverwriteExisting: (int)swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd);
            }
            else
            {
                // Creates a new custom property
                customPropertyManager.Add3(
                    FieldName: CustomPropertyName,
                    FieldType: 3,
                    FieldValue: NewCustomPropertyValue.ToString(),
                    OverwriteExisting: (int)swCustomPropertyAddOption_e.swCustomPropertyOnlyIfNew);
            }
        }


        /// <summary>
        /// Looks for a custom Yes or No property, if there is one, changes its value. If the custom Yes or No property
        /// does not exist, creates a new one.
        /// </summary>
        /// <param name="SolidWorksApplication"></param>
        /// <param name="ModelDocument"></param>
        public static void AddOrEditCustomYesNoProperty(ModelDoc2 ModelDocument, string CustomPropertyName, YesNo NewCustomPropertyValue)
        {
            //Get the custom property manager
            CustomPropertyManager customPropertyManager = ModelDocument.Extension.CustomPropertyManager[""];

            // Store the property type
            int propertyType = customPropertyManager.GetType2(CustomPropertyName);

            /* Gets the value and the evaluated value of the specified custom property.
             Result codes when getting custom properties: 0 = Cached value was returned,
            1 = Custom property does not exist, 2 = Resolved value was returned.*/
            int status = customPropertyManager.Get6(
                FieldName: CustomPropertyName,
                UseCached: false,
                ValOut: out _,
                ResolvedValOut: out _,
                WasResolved: out _,
                LinkToProperty: out _);

            // If the custom property exists
            if (status == 2 && propertyType == 11)
            {
                // Change the value of the custom property
                customPropertyManager.Add3(
                    FieldName: CustomPropertyName,
                    FieldType: 11,
                    FieldValue: NewCustomPropertyValue.ToString(),
                    OverwriteExisting: (int)swCustomPropertyAddOption_e.swCustomPropertyReplaceValue);
            }
            // If the custom property exists, but its type is not 'text'
            else if (status == 2)
            {
                // Add a new custom property of a type 'text' and delete previously found one.
                customPropertyManager.Add3(
                    FieldName: CustomPropertyName,
                    FieldType: 11,
                    FieldValue: NewCustomPropertyValue.ToString(),
                    OverwriteExisting: (int)swCustomPropertyAddOption_e.swCustomPropertyDeleteAndAdd);
            }
            else
            {
                // Creates a new custom property
                customPropertyManager.Add3(
                    FieldName: CustomPropertyName,
                    FieldType: 11,
                    FieldValue: NewCustomPropertyValue.ToString(),
                    OverwriteExisting: (int)swCustomPropertyAddOption_e.swCustomPropertyOnlyIfNew);
            }
        }


        /// <summary>
        /// Gets the custom property value if the type of the property is text
        /// </summary>
        /// <param name="ModelDocument"></param>
        /// <param name="CustomPropertyName"></param>
        /// <param name="Value"></param>
        /// <returns></returns>
        public static bool GetCustomTextPropertyValue(ModelDoc2 ModelDocument, string CustomPropertyName, out string Value)
        {
            //Get the custom property manager
            CustomPropertyManager customPropertyManager = ModelDocument.Extension.CustomPropertyManager[""];

            // Check if there is a custom property with the specified name
            int status = customPropertyManager.Get6(
                FieldName: CustomPropertyName,
                UseCached: false,
                ValOut: out Value,
                ResolvedValOut: out _,
                WasResolved: out _,
                LinkToProperty: out _);

            // Store the property type
            int propertyType = customPropertyManager.GetType2(CustomPropertyName);

            // Return true if the property with a specific name was found AND if the type of the property is text.
            return status != 0 && propertyType == 30;
        }


        /// <summary>
        /// Gets the custom property value if the type of the property is number
        /// </summary>
        /// <param name="ModelDocument"></param>
        /// <param name="CustomPropertyName"></param>
        /// <param name="Value"></param>
        /// <returns></returns>
        public static bool GetCustomNumberPropertyValue(ModelDoc2 ModelDocument, string CustomPropertyName, out double Value)
        {
            Value = 0;

            //Get the custom property manager
            CustomPropertyManager customPropertyManager = ModelDocument.Extension.CustomPropertyManager[""];

            // Check if there is a custom property with the specified name
            int status = customPropertyManager.Get6(
                FieldName: CustomPropertyName,
                UseCached: false,
                ValOut: out string customPropertyValueString,
                ResolvedValOut: out _,
                WasResolved: out _,
                LinkToProperty: out _);

            // Store the property type
            int propertyType = customPropertyManager.GetType2(CustomPropertyName);

            // If custom property type is number
            if (propertyType == 3)
            {
                // Convert the value of the custom property from string to double
                double.TryParse(customPropertyValueString, out Value);
            };

            // Return true if the property with a specific name was found AND if the type of the property is text.
            return status != 0 && propertyType == 3;
        }


        /// <summary>
        /// Gets the custom property value if the type of the property is Yes or No
        /// </summary>
        /// <param name="ModelDocument"></param>
        /// <param name="CustomPropertyName"></param>
        /// <param name="Value"></param>
        /// <returns></returns>
        public static bool GetCustomYesNoPropertyValue(ModelDoc2 ModelDocument, string CustomPropertyName, out YesNo Value)
        {
            // Set default value (which is Yes)
            Value = default;

            //Get the custom property manager
            CustomPropertyManager customPropertyManager = ModelDocument.Extension.CustomPropertyManager[""];

            // Check if there is a custom property with the specified name
            int status = customPropertyManager.Get6(
                FieldName: CustomPropertyName,
                UseCached: false,
                ValOut: out string customPropertyValueString,
                ResolvedValOut: out _,
                WasResolved: out _,
                LinkToProperty: out _);

            // Store the property type
            int propertyType = customPropertyManager.GetType2(CustomPropertyName);

            // If custom property type is Yes or No
            if (propertyType == 11)
            {
                // Convert the value of the custom property from string to Enum (Yes, No)
                Enum.TryParse(customPropertyValueString, out Value);
            };

            // Return true if the property with a specific name was found AND if the type of the property is text.
            return status != 0 && propertyType == 11;
        }


        /// <summary>
        /// Checks if a provided Entity represents a SOLIDWORKS datum plane.
        /// </summary>
        /// <param name="Entity">The Entity object to be checked.</param>
        /// <returns>True if the Entity is a SOLIDWORKS datum plane; otherwise, false.</returns>
        public static bool IsEntityPlane(Entity Entity) => (Entity.GetType() == (int)swSelectType_e.swSelDATUMPLANES);


        /// <summary>
        /// Determines if a given PlaneFeature represents a major plane in the specified document.
        /// </summary>
        /// <param name="PlaneFeature">The PlaneFeature to evaluate.</param>
        /// <param name="Document">The ModelDoc2 where the plane is located.</param>
        /// <param name="PlaneType">The type of major plane to compare against (e.g., MajorPlane.FrontPlane).</param>
        /// <returns>True if the PlaneFeature represents the specified major plane, false otherwise.</returns>
        public static bool IsEntityMajorPlane(Feature PlaneFeature, ModelDoc2 Document, MajorPlane PlaneType) => (PlaneFeature.Name == Utilities.GetMajorPlane(Document, PlaneType).Name);


        /// <summary>
        /// Determines if a given PlaneFeature represents a major plane in the specified component's document.
        /// </summary>
        /// <param name="PlaneFeature">The PlaneFeature to evaluate.</param>
        /// <param name="Document">The Component2 containing the plane.</param>
        /// <param name="PlaneType">The type of major plane to compare against (e.g., MajorPlane.FrontPlane).</param>
        /// <returns>True if the PlaneFeature represents the specified major plane, false otherwise.</returns>
        public static bool IsEntityMajorPlane(Feature PlaneFeature, Component2 Document, MajorPlane PlaneType) => (PlaneFeature.Name == Utilities.GetMajorPlane(Document, PlaneType).Name);

        public static bool IsEntityAxis(Entity Entity) => (Entity.GetType() == (int)swSelectType_e.swSelDATUMAXES);


        /// <summary>
        /// Checks if a given entity belongs to a specific component by comparing names.
        /// </summary>
        /// <param name="Entity">The entity to check.</param>
        /// <param name="Component">The component to check against.</param>
        /// <returns>True if the entity belongs to the specified component; otherwise, false.</returns>
        public static bool IsEntityInSpecificComponent(Entity Entity, Component2 Component) => IsEntityInComponent(Entity) && Entity.GetComponent().Name2 == Component.Name2;


        /// <summary>
        /// Checks if an entity belongs to the component.
        /// </summary>
        /// <param name="Entity">The entity to be inspected.</param>
        /// <returns>True if the entity belongs to the component; false if entity belongs to the assembly.</returns>
        public static bool IsEntityInComponent(Entity Entity) => (Entity.GetComponent() != null);


        /// <summary>
        /// Applies a mathematical transformation to a vector expressed as 3D coordinates
        /// </summary>
        /// <param name="transform"></param>
        /// <param name="Vector"></param>
        /// <returns></returns>
        public static double[] TransformVector(ISldWorks SolidWorksApplication, MathTransform transform, double[] Vector)
        {
            //Create a SolidWorks MathVector object from the input vector
            //.
            MathVector mathVector = SolidWorksApplication.GetMathUtility().CreateVector(Vector);

            //Apply the provided mathematical transformation to the vector. The 'MultiplyTransform'
            //method updates the MathVector with the results of the transformation.
            mathVector = mathVector.MultiplyTransform(transform);

            //Return the transformed vector as a double array.
            return mathVector.ArrayData;
        }


        /// <summary>
        /// Determines if a given plane in SolidWorks is parallel to a specified major plane (Front, Top, or Right).
        /// </summary>
        /// <param name="SolidWorksApplication">The current SolidWorks application instance.</param>
        /// <param name="ModelDocument">The SolidWorks model document.</param>
        /// <param name="Plane">The SolidWorks Feature representing the plane to check.</param>
        /// <param name="PlaneType">The type of major plane to compare against (Front, Top, or Right).</param>
        /// <returns>True if the plane is parallel to the specified major plane, false otherwise.</returns>
        public static bool IsPlaneParallelToMajorPlane(ISldWorks SolidWorksApplication, Feature Plane, MajorPlane PlaneType)
        {
            // Obtain the Reference Plane
            RefPlane refPlane = (RefPlane)Plane.GetSpecificFeature2();

            // Get the plane's transform (position and orientation)
            MathTransform transform = refPlane.Transform;

            // Normal vector of a plane is perpendicular to it. Here, we check if the transformed
            // Z-axis aligns with a major axis after rounding (allowing for floating-point tolerance).
            double[] transformedNormal = TransformVector(SolidWorksApplication, transform, new double[3] { 0, 0, 1 });

            double roundedXCoordinate = Math.Round(transformedNormal[0], 13);
            double roundedYCoordinate = Math.Round(transformedNormal[1], 13);
            double roundedZCoordinate = Math.Round(transformedNormal[2], 13);

            switch (PlaneType)
            {
                case MajorPlane.Front:
                    return (roundedXCoordinate == 0 && roundedYCoordinate == 0 && (roundedZCoordinate == 1 || roundedZCoordinate == -1));
                case MajorPlane.Top:
                    return (roundedXCoordinate == 0 && (roundedYCoordinate == 1 || roundedYCoordinate == -1) && roundedZCoordinate == 0);
                default: // Plane is right
                    return ((roundedXCoordinate == 1 || roundedXCoordinate == -1) && roundedYCoordinate == 0 && roundedZCoordinate == 0);
            }
        }


        /// <summary>
        /// Extracts top-level components from a SolidWorks assembly document.
        /// </summary>
        /// <param name="Assembly">The SolidWorks assembly document (ModelDoc2) to process.</param>
        /// <returns>A List where with Component2 objects.</returns>
        public static List<Component2> GetTopLevelComponents(ModelDoc2 Assembly)
        {
            List<Component2> returnCollection = new List<Component2>();

            int componentCounter = 1;
            Feature loopComponent = Utilities.GetNthFeatureOfType(Assembly, FeatureType.Component, componentCounter);

            while (loopComponent != null)
            {
                returnCollection.Add((Component2)loopComponent.GetSpecificFeature2());
                componentCounter++;
                loopComponent = Utilities.GetNthFeatureOfType(Assembly, FeatureType.Component, componentCounter);
            }

            return returnCollection;
        }
    }
}
