using SolidWorks.Interop.sldworks;

namespace AddinWithTaskpane
{
    internal class EventHandlers
    {
        // Have access to SW application
        static SldWorks swApp;

        public void AssignSWApp(SldWorks inputswApp)
        {
            swApp = inputswApp;
        }


        /// <summary>
        /// Event that notifies when the document is changed
        /// </summary>
        private static DSldWorksEvents_ActiveModelDocChangeNotifyEventHandler documentChanged = 
            new DSldWorksEvents_ActiveModelDocChangeNotifyEventHandler(OnModelDocChanged);

        /// <summary>
        /// A method to attach application level event handlers. This needs to be called from the main ConnectToSW() function.
        /// </summary>
        /// <param name="swApp"></param>
        public static void AttachEventHandlers(SldWorks swApp)
        {
            swApp.ActiveModelDocChangeNotify += documentChanged;
        }

        /// <summary>
        /// A method to detach application level event handlers. This needs to be called from the main DisconnectFromSW() function.
        /// </summary>
        /// <param name="swApp"></param>
        public static void DetachEventHandlers(SldWorks swApp)
        {
            swApp.ActiveModelDocChangeNotify -= documentChanged;
        }

        /// <summary>
        /// Functionality that is required when the document has been changed.
        /// </summary>
        /// <returns></returns>
        private static int OnModelDocChanged()
        {
            ModelDoc2 assembly = (ModelDoc2)swApp.ActiveDoc;

            string customPropertyName = "Tipas";

            //Get the custom property manager
            CustomPropertyManager customPropertyManager = assembly.Extension.CustomPropertyManager[""];

            // Check if there is a custom property with the specified name
            int status = customPropertyManager.Get6(
                FieldName: customPropertyName,
                UseCached: false,
                ValOut: out string customPropertyValue,
                ResolvedValOut: out _,
                WasResolved: out _,
                LinkToProperty: out _);

            // Store the property type
            int propertyType = customPropertyManager.GetType2(customPropertyName);

            return 0;
        }
    }
}
