using System;
using System.Resources;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace RoomInfoOutlookAddIn
{
    partial class OccupancyFormRegion
    {
        ResourceManager _resourceManager;
        Outlook.AppointmentItem _outlookAppointmentItem;
        bool _initializing;
        //IEventService _eventService;
        #region Formularbereichsfactory 

        [Microsoft.Office.Tools.Outlook.FormRegionMessageClass(Microsoft.Office.Tools.Outlook.FormRegionMessageClassAttribute.Appointment)]
        [Microsoft.Office.Tools.Outlook.FormRegionName("RoomInfoOutlookAddIn.OccupancyFormRegion")]
        public partial class OccupancyFormRegionFactory
        {
            // Tritt ein, bevor der Formularbereich initialisiert wird.
            // Um die Anzeige des Formularbereichs zu verhindern, legen Sie "e.Cancel" auf "true" fest.
            // Verwenden Sie e.OutlookItem, um einen Verweis auf das aktuelle Outlook-Element abzurufen.
            private void OccupancyFormRegionFactory_FormRegionInitializing(object sender, Microsoft.Office.Tools.Outlook.FormRegionInitializingEventArgs e)
            {
                Outlook.AppointmentItem outlookAppointmentItem = (Outlook.AppointmentItem)e.OutlookItem;
                if (outlookAppointmentItem != null && outlookAppointmentItem.UserProperties.Find("Occupancy") != null) return;
                e.Cancel = true;
            }
        }

        #endregion

        // Tritt ein, bevor der Formularbereich angezeigt wird.
        // Verwenden Sie this.OutlookItem, um einen Verweis auf das aktuelle Outlook-Element abzurufen.
        // Verwenden Sie this.OutlookFormRegion, um einen Verweis auf den Formularbereich abzurufen.
        private void OccupancyFormRegion_FormRegionShowing(object sender, EventArgs e)
        {
            //_eventService = ServiceLocator.Current.GetInstance<IEventService>();
            _resourceManager = Properties.LanguageResources.ResourceManager;
            occupancyComboBox.Items.Add(_resourceManager.GetString("DropDown_ItemLabel_Occupancy_Free"));
            occupancyComboBox.Items.Add(_resourceManager.GetString("DropDown_ItemLabel_Occupancy_Present"));
            occupancyComboBox.Items.Add(_resourceManager.GetString("DropDown_ItemLabel_Occupancy_Absent"));
            occupancyComboBox.Items.Add(_resourceManager.GetString("DropDown_ItemLabel_Occupancy_Busy"));
            occupancyComboBox.Items.Add(_resourceManager.GetString("DropDown_ItemLabel_Occupancy_Occupied"));
            occupancyComboBox.Items.Add(_resourceManager.GetString("DropDown_ItemLabel_Occupancy_Locked"));
            occupancyComboBox.Items.Add(_resourceManager.GetString("DropDown_ItemLabel_Occupancy_Home"));
            var formRegionControl = sender as Microsoft.Office.Tools.Outlook.FormRegionControl;
            _outlookAppointmentItem = (Outlook.AppointmentItem)formRegionControl.OutlookItem;
            if (_outlookAppointmentItem != null && _outlookAppointmentItem.UserProperties.Find("Occupancy") != null)
            {
                _initializing = true;
                occupancyComboBox.SelectedIndex = _outlookAppointmentItem.UserProperties.Find("Occupancy").Value;
            }
        }

        // Tritt ein, wenn der Formularbereich geschlossen wird.
        // Verwenden Sie this.OutlookItem, um einen Verweis auf das aktuelle Outlook-Element abzurufen.
        // Verwenden Sie this.OutlookFormRegion, um einen Verweis auf den Formularbereich abzurufen.
        private void OccupancyFormRegion_FormRegionClosed(object sender, EventArgs e)
        {
        }

        private void occupancyComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            _outlookAppointmentItem.UserProperties.Find("Occupancy").Value = (sender as ComboBox).SelectedIndex;
            if (!_initializing) _outlookAppointmentItem.Save();
            _initializing = false;
        }
    }
}
