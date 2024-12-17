using McTools.Xrm.Connection;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.Services.Description;
using System.Windows.Forms;
using XrmToolBox.Extensibility;

namespace Excel_Importer
{
    public partial class MyPluginControl : PluginControlBase
    {
        private Settings mySettings;
        private IOrganizationService service;  // This will store the service connection

        public MyPluginControl()
        {
            InitializeComponent();
        }

        private void MyPluginControl_Load(object sender, EventArgs e)
        {
            //ShowInfoNotification("This is a notification that can lead to XrmToolBox repository", new Uri("https://github.com/MscrmTools/XrmToolBox"));

            // Loads or creates the settings for the plugin
            if (!SettingsManager.Instance.TryLoad(GetType(), out mySettings))
            {
                mySettings = new Settings();

                LogWarning("Settings not found => a new settings file has been created!");
            }
            else
            {
                LogInfo("Settings found and loaded");
            }
            LoadEntities();
            progressBar1.Visible = false;
            label6.Visible = true;
            button1.Visible = false;
            button3.Visible = false;
            label8.Visible = true;
            panel1.BringToFront();

        }

        private void tsbClose_Click(object sender, EventArgs e)
        {
            CloseTool();
        }

        private void tsbSample_Click(object sender, EventArgs e)
        {
            // The ExecuteMethod method handles connecting to an
            // organization if XrmToolBox is not yet connected
            ExecuteMethod(GetAccounts);
        }

        private void GetAccounts()
        {
            WorkAsync(new WorkAsyncInfo
            {
                Message = "Getting accounts",
                Work = (worker, args) =>
                {
                    args.Result = Service.RetrieveMultiple(new QueryExpression("account")
                    {
                        TopCount = 50
                    });
                },
                PostWorkCallBack = (args) =>
                {
                    if (args.Error != null)
                    {
                        MessageBox.Show(args.Error.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    var result = args.Result as EntityCollection;
                    if (result != null)
                    {
                        MessageBox.Show($"Found {result.Entities.Count} accounts");
                    }
                }
            });
        }

        /// <summary>
        /// This event occurs when the plugin is closed
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void MyPluginControl_OnCloseTool(object sender, EventArgs e)
        {
            // Before leaving, save the settings
            SettingsManager.Instance.Save(GetType(), mySettings);
        }

        /// <summary>
        /// This event occurs when the connection has been updated in XrmToolBox
        /// </summary>
        public override void UpdateConnection(IOrganizationService newService, ConnectionDetail detail, string actionName, object parameter)
        {
            base.UpdateConnection(newService, detail, actionName, parameter);
            service = newService;

            if (mySettings != null && detail != null)
            {
                mySettings.LastUsedOrganizationWebappUrl = detail.WebApplicationUrl;
                LogInfo("Connection has changed to: {0}", detail.WebApplicationUrl);
            }

            //LoadEntities();
        }

        private void ImportExcel(DataGridView dataGridView, TextBox textBox, ComboBox toolStripComboBox1)
        {
            try
            {
                // Set the license context for EPPlus
                OfficeOpenXml.ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                using (var openFileDialog = new OpenFileDialog())
                {
                    openFileDialog.Filter = "Excel Files|*.xlsx;*.xls";
                    if (openFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        var filePath = openFileDialog.FileName;

                        using (var package = new ExcelPackage(new FileInfo(filePath)))
                        {
                            var worksheet = package.Workbook.Worksheets[0];
                            var dataTable = new DataTable();

                            // Clear ComboBox items
                            toolStripComboBox1.Items.Clear();

                            // Load header row
                            for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                            {
                                var columnName = worksheet.Cells[1, col].Text;
                                dataTable.Columns.Add(columnName);
                                toolStripComboBox1.Items.Add(columnName);
                            }

                            // Set default selection if ComboBox has items
                            if (toolStripComboBox1.Items.Count > 0)
                            {
                                toolStripComboBox1.SelectedIndex = 0; // Select the first item by default
                            }

                            // Load data rows
                            for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
                            {
                                var newRow = dataTable.NewRow();
                                bool isEmptyRow = true;

                                for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                                {
                                    var cellValue = worksheet.Cells[row, col].Text;

                                    // Check if the cell contains any non-empty value
                                    if (!string.IsNullOrEmpty(cellValue))
                                    {
                                        isEmptyRow = false;
                                    }

                                    newRow[col - 1] = cellValue;
                                }

                                // Only add the row if it's not empty
                                if (!isEmptyRow)
                                {
                                    dataTable.Rows.Add(newRow);
                                }
                            }

                            // Add a checkbox column
                            var checkBoxColumn = new DataGridViewCheckBoxColumn
                            {
                                HeaderText = "Select",
                                Name = "Select",
                                Width = 50
                            };

                            if (!dataGridView.Columns.Contains("Select"))
                            {
                                dataGridView.Columns.Add(checkBoxColumn);
                            }

                            // Disable the "new row" feature
                            dataGridView.AllowUserToAddRows = false;

                            // Bind to DataGridView
                            dataGridView.DataSource = dataTable;

                            // Total record line
                            var totalRecord = dataTable.Rows.Count;

                            // Show row count
                            textBox.Text = totalRecord.ToString();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred while importing the file: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }



        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            ImportExcel(dataGridView1, textBox1, comboBox4);
            textBox2.Text = " ";
            if(textBox2.Text == " ")
            {
                button1.Visible = false;
            }
            checkBox1.Checked = false;
        }

        
        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            //var SelectedCount = 0;
            if (e.ColumnIndex == dataGridView1.Columns["Select"].Index && e.RowIndex >= 0)
            {
                DataGridViewCheckBoxCell checkBoxCell = (DataGridViewCheckBoxCell)dataGridView1.Rows[e.RowIndex].Cells["Select"];

                // Toggle the checkbox value
                bool isChecked = (checkBoxCell.Value == null ? false : (bool)checkBoxCell.Value);
                bool checkorNot = (bool)(checkBoxCell.Value = !isChecked);

                if(!checkorNot)
                {
                    checkBox1.Checked = false;
                }
                else
                {
                    CheckAllCheckboxes();
                }

                // Update the selected count
                UpdateSelectedCount();
            }

        }

        // Method to check if all checkboxes are checked
        private void CheckAllCheckboxes()
        {
            bool allChecked = true;

            // Loop through each row in the DataGridView
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                // Check if the row contains a checkbox in the "Select" column
                if (row.Cells["Select"] is DataGridViewCheckBoxCell checkBoxCell)
                {
                    // If any checkbox is unchecked, set allChecked to false
                    if (checkBoxCell.Value == null || !(bool)checkBoxCell.Value)
                    {
                        allChecked = false;
                        break; // No need to continue checking
                    }
                }
            }

            // If all checkboxes are checked, set checkBox1 to checked
            checkBox1.Checked = allChecked;
        }



        //private void checkBox1_CheckedChanged(object sender, EventArgs e)
        //{

        //}

        private void checkBox1_Click(object sender, EventArgs e)
        {
            // Check if the checkbox is checked
            bool selectAll = checkBox1.Checked;

            //Loop through each row in the DataGridView
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                //Update the checkBox cell in the 'Select' column
                DataGridViewCheckBoxCell checkBoxCell = row.Cells["Select"] as DataGridViewCheckBoxCell;
                if (checkBoxCell != null)
                {
                    checkBoxCell.Value = selectAll;
                }
            }

            // Update the selected count
            UpdateSelectedCount();
        }


        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            //var countSelected = 0;
            //if(toolStripComboBox1.SelectedItem == null)
            //{
            //    MessageBox.Show("Please select a column from the dropdown!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            //    return;
            //}
            if(comboBox4.SelectedItem == null)
            {
                MessageBox.Show("Please select a column from the dropdown!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            //Get the selected column name from the ComboBox
            //string selectedColumn = toolStripComboBox1.SelectedItem.ToString();
            string selectedColumn = comboBox4.SelectedItem.ToString();

            //List to store the records to deactivate 
            List<string> recordsToDeactivate = new List<string>();

            //Loop through DataGridView rows
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                // Get the checkbox cell value

                // Check if the checkbox in the "Select" column is checked
                var checkBoxCell = row.Cells["Select"] as DataGridViewCheckBoxCell;
                if (checkBoxCell != null && checkBoxCell.Value != null && (bool)checkBoxCell.Value) 
                {
                    //countSelected++;
                    //label8.Text = countSelected.ToString();

                    // Retrieve the value from the selected column
                    var cellValue = row.Cells[selectedColumn]?.Value?.ToString();
                    if (!string.IsNullOrEmpty(cellValue)) 
                    {
                        recordsToDeactivate.Add(cellValue);
                    }
                }
            }

            // check if any records are selected
            if (recordsToDeactivate.Count == 0)
            {
                MessageBox.Show("No records selected for deactivation.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // Perform the deactivation logic
            DeactivateRecords(recordsToDeactivate);
        }

        //private void DeactivateRecords(List<string> records)
        //{
        //    string selectedEntityLogicalName = null;
        //    richTextBox1.Clear();
        //    int errorCount = 0; // Counter to track the number of errors
        //    textBox3.Text = " ";
        //    int successCount = 0;

        //    if (comboBox1.SelectedItem != null)
        //    {
        //        // Directly retrieve the logical name from the selected item
        //        selectedEntityLogicalName = comboBox1.SelectedItem.ToString();
        //    }



        //    // Set the progress bar maximum value to the total number of records
        //    progressBar1.Maximum = records.Count;
        //    progressBar1.Value = 0;

        //    // Hide the progress bar initially (if you want to start without showing it)
        //    progressBar1.Visible = false;

        //    //Initialize the progress counter
        //    label6.Text = "0/" + records.Count;

        //    try
        //    {
        //        if (comboBox2.SelectedItem == null)
        //        {
        //            MessageBox.Show("Please Select Status !!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        //        }
        //        else if (comboBox3.SelectedItem == null)
        //        {
        //            MessageBox.Show("Please Select Status Reason!!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        //        }
        //        else
        //        {
        //            // Get the selected statecode and statuscode from the dropdowns
        //            var selectedStateCode = comboBox2.SelectedItem.ToString();
        //            var selectedStatusCode = comboBox3.SelectedItem.ToString();

        //            int stateCodeValue = int.Parse(selectedStateCode.Split(' ')[0]); // Extract the numeric value from the selected statecode
        //            int statusCodeValue = int.Parse(selectedStatusCode.Split(' ')[0]); // Extract the numeric value from the selected 

        //            // Show the progress bar when starting the process
        //            progressBar1.Visible = true;

        //            //show progress counter
        //            label6.Visible = true;

        //            // Display the records to deactivate
        //            //Iterate through each record ID
        //            foreach (var recordId in records)
        //            {
        //                // Create a request to set the state and status
        //                var setStateRequest = new Microsoft.Crm.Sdk.Messages.SetStateRequest
        //                {
        //                    EntityMoniker = new EntityReference(selectedEntityLogicalName, Guid.Parse(recordId)),
        //                    //State = new OptionSetValue(1),  // 1 usually corresponds to "Inactive"
        //                    //Status = new OptionSetValue(2)  // Use the specific Status code for "Deactivated"
        //                    State = new OptionSetValue(stateCodeValue),
        //                    Status = new OptionSetValue(statusCodeValue)
        //                };

        //                //Execute the request
        //                Service.Execute(setStateRequest);

        //                // Increment the success counter
        //                successCount++;
        //                progressBar1.Value = successCount;

        //                //Update the progress label with the current success count(text format: "X/total")
        //                label6.Text = successCount + "/" + records.Count;

        //                // Refresh UI
        //                Application.DoEvents();

        //            }

        //            MessageBox.Show("Status changed successfully!!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
        //            textBox2.Text = successCount.ToString();

        //            //success count textbox
        //            if (textBox2.Text != null)
        //            {
        //                progressBar1.Visible = false;
        //                label6.Visible = false;
        //                button1.Visible = true;
        //            }

        //        }


        //    }
        //    catch (Exception ex)
        //    {
        //        //MessageBox.Show($"An error occurred while deactivating records: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        //        errorCount++;
        //        textBox3.Text = errorCount.ToString();
        //        string errorMessage = $"Error at {DateTime.Now}: {ex.Message}\n{ex.StackTrace}\n";

        //        // Append the error message to the Log Tab
        //        richTextBox1.AppendText(errorMessage);

        //        //Error count textbox
        //        if (textBox3.Text != null)
        //        {
        //            progressBar1.Visible = false;
        //            label6.Visible = false;
        //            button1.Visible = false;
        //            textBox2.Text = " ";
        //        }
        //    }
        //}

        //private void DeactivateRecords(List<string> records)
        //{
        //    string selectedEntityLogicalName = null;
        //    richTextBox1.Clear();
        //    int errorCount = 0;
        //    textBox3.Text = " ";
        //    int successCount = 0;

        //    if (comboBox1.SelectedItem != null)
        //    {
        //        selectedEntityLogicalName = comboBox1.SelectedItem.ToString();
        //    }

        //    progressBar1.Maximum = records.Count;
        //    progressBar1.Value = 0;
        //    progressBar1.Visible = false;
        //    label6.Text = "0/" + records.Count;

        //    if (comboBox2.SelectedItem == null)
        //    {
        //        MessageBox.Show("Please Select Status !!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        //        return;
        //    }

        //    if (comboBox3.SelectedItem == null)
        //    {
        //        MessageBox.Show("Please Select Status Reason!!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        //        return;
        //    }

        //    var selectedStateCode = comboBox2.SelectedItem.ToString();
        //    var selectedStatusCode = comboBox3.SelectedItem.ToString();

        //    int stateCodeValue = int.Parse(selectedStateCode.Split(' ')[0]);
        //    int statusCodeValue = int.Parse(selectedStatusCode.Split(' ')[0]);

        //    progressBar1.Visible = true;
        //    label6.Visible = true;

        //    foreach (var recordId in records)
        //    {
        //        try
        //        {
        //            var setStateRequest = new Microsoft.Crm.Sdk.Messages.SetStateRequest
        //            {
        //                EntityMoniker = new EntityReference(selectedEntityLogicalName, Guid.Parse(recordId)),
        //                State = new OptionSetValue(stateCodeValue),
        //                Status = new OptionSetValue(statusCodeValue)
        //            };

        //            Service.Execute(setStateRequest);

        //            successCount++;
        //            progressBar1.Value = successCount;
        //            label6.Text = successCount + "/" + records.Count;

        //            Application.DoEvents();
        //        }
        //        catch (Exception ex)
        //        {
        //            errorCount++;
        //            progressBar1.Visible = true;
        //            progressBar1.Value = errorCount;
        //            label6.Text = errorCount + "/" + records.Count;
        //            Application.DoEvents();
        //            string errorMessage = $"Error processing record {recordId}: {ex.Message}\n{ex.StackTrace}\n";
        //            richTextBox1.AppendText(errorMessage);

        //        }
        //    }

        //    // Update success and error counts
        //    textBox2.Text = successCount.ToString();
        //    textBox3.Text = errorCount.ToString();

        //    progressBar1.Visible = false;
        //    label6.Visible = false;

        //    if (errorCount > 0)
        //    {
        //        MessageBox.Show($"{errorCount} record(s) failed to deactivate. Check the error log for details.", "Partial Success", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        //        //Error count textbox
        //        if (textBox3.Text != null)
        //        {
        //            progressBar1.Visible = false;
        //            label6.Visible = false;
        //            button1.Visible = false;
        //            textBox2.Text = "0";
        //        }
        //    }
        //    else
        //    {
        //        MessageBox.Show("Status changed successfully for all records!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);

        //        //success count textbox
        //        if (textBox2.Text != null)
        //        {
        //            progressBar1.Visible = false;
        //            label6.Visible = false;
        //            button1.Visible = true;
        //        }
        //    }
        //}


        private void DeactivateRecords(List<string> records)
        {
            string selectedEntityLogicalName = null;
            richTextBox1.Clear();
            int errorCount = 0;
            int successCount = 0;

            textBox3.Text = string.Empty; // Error count text box
            textBox2.Text = string.Empty; // Success count text box

            List<string> successfulRecordIds = new List<string>(); //List to store successful records.

            // Validate entity selection
            if (comboBox1.SelectedItem != null)
            {
                selectedEntityLogicalName = comboBox1.SelectedItem.ToString();
            }
            else
            {
                MessageBox.Show("Please select an entity!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // Validate state and status codes
            if (comboBox2.SelectedItem == null || comboBox3.SelectedItem == null)
            {
                MessageBox.Show("Please select both Status and Status Reason!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // Extract state and status codes
            var stateCodeValue = int.Parse(comboBox2.SelectedItem.ToString().Split(' ')[0]);
            var statusCodeValue = int.Parse(comboBox3.SelectedItem.ToString().Split(' ')[0]);

            // Initialize progress bar and label
            progressBar1.Maximum = records.Count;
            progressBar1.Value = 0;
            progressBar1.Visible = true;
            label6.Visible = true;
            label6.Text = $"0/{records.Count}";

            // Process records in batches
            const int batchSize = 50; // Adjust batch size as needed
            for (int i = 0; i < records.Count; i += batchSize)
            {
                var batch = records.Skip(i).Take(batchSize).ToList();
                foreach (var recordId in batch)
                {
                    try
                    {
                        // Execute SetStateRequest for each record
                        var setStateRequest = new Microsoft.Crm.Sdk.Messages.SetStateRequest
                        {
                            EntityMoniker = new EntityReference(selectedEntityLogicalName, Guid.Parse(recordId)),
                            State = new OptionSetValue(stateCodeValue),
                            Status = new OptionSetValue(statusCodeValue)
                        };
                        Service.Execute(setStateRequest);

                        // Increment success count
                        successCount++;
                        successfulRecordIds.Add(recordId);
                       
                    }
                    catch (Exception ex)
                    {
                        // Increment error count and log errors
                        errorCount++;
                        string errorMessage = $"Error processing record {recordId}: {ex.Message}\n";
                        richTextBox1.AppendText(errorMessage);
                        // Allow the UI to refresh
                        Application.DoEvents();
                    }

                    // Update progress bar and label
                    progressBar1.Value = successCount + errorCount;
                    label6.Text = $"{successCount + errorCount}/{records.Count}";

                    // Allow the UI to refresh
                    Application.DoEvents();
                }

                // Allow the UI to refresh
                Application.DoEvents();
            }

            // Display final results
            textBox2.Text = successCount.ToString(); // Success count
            textBox3.Text = errorCount.ToString();   // Error count

            progressBar1.Visible = false;
            label6.Visible = false;

            if(textBox2.Text != null)
            {
                button1.Visible = true;
                button1.Tag = successfulRecordIds;
            }
            else
            {
                button1.Visible = false;
            }

            if (errorCount > 0)
            {
                MessageBox.Show($"{errorCount} record(s) failed. Check the error log for details.", "Partial Success", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                button3.Visible = true;
            }
            else
            {
                MessageBox.Show("All records processed successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                button3.Visible = false;
            }
        }



        private void LoadEntities()
        {
            try
            {
                // Display loading message while retrieving entities
                WorkAsync(new WorkAsyncInfo
                {
                    Message = "Retrieving entities...",
                    Work = (worker, args) =>
                    {
                        // Define request to retrieve entity metadata
                        var retrieveAllEntitiesRequest = new RetrieveAllEntitiesRequest
                        {
                            EntityFilters = EntityFilters.Entity,
                            RetrieveAsIfPublished = true
                        };

                        // Execute the request
                        var response = (RetrieveAllEntitiesResponse)Service.Execute(retrieveAllEntitiesRequest);

                        // Process and store the result
                        args.Result = response.EntityMetadata
                            .OrderBy(entity => entity.LogicalName)
                            .Select(entity => entity.LogicalName)
                            .ToList();
                    },
                    PostWorkCallBack = args =>
                    {
                        // Handle errors
                        if (args.Error != null)
                        {
                            MessageBox.Show($"Error: {args.Error.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }

                        // Handle successful response
                        var entities = args.Result as System.Collections.Generic.List<string>;
                        if (entities != null)
                        {
                            // Bind entities to the dropdown
                            comboBox1.DataSource = entities;
                            //MessageBox.Show($"Successfully retrieved {entities.Count} entities.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                        else
                        {
                            MessageBox.Show("No entities were retrieved.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }
                });
            }
            catch (Exception ex)
            {
                // Handle unexpected exceptions
                MessageBox.Show($"An error occurred: {ex.Message}", "Critical Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            comboBox2.ResetText();
            comboBox3.ResetText();
            if (comboBox1.SelectedItem != null) 
            {
                string selectedEntity = comboBox1.SelectedItem.ToString();
                LoadStateCodes(selectedEntity);
            }
        }
        
        private void LoadStateCodes(string entityLogicalName)
        {
            try
            {
                // Clear the dropdown before adding new items
                comboBox2.Items.Clear();

                var entityMetadataRequest = new RetrieveEntityRequest
                {
                    EntityFilters = EntityFilters.Attributes,
                    LogicalName = entityLogicalName 
                };

                // Retrieve metadata for the selected entity.
                RetrieveEntityResponse entityMetadataResponse = (RetrieveEntityResponse)service.Execute(entityMetadataRequest);

                // Find the statecode attribute in the entity metadata.
                var statecodeAttributeMetadata = entityMetadataResponse.EntityMetadata.Attributes
                    .FirstOrDefault(a => a.LogicalName == "statecode");

                if (statecodeAttributeMetadata != null && statecodeAttributeMetadata is EnumAttributeMetadata enumAttributeMetadata)
                {
                    // Retrieve the valid statecode options (Active, Inactive, etc.)
                    var statecodeOptions = enumAttributeMetadata.OptionSet.Options;

                    foreach (var option in statecodeOptions)
                    {
                        comboBox2.Items.Add($"{option.Value} {option.Label.UserLocalizedLabel.Label}");
                    }
                }
                else
                {
                    Console.WriteLine("Statecode field is not available or doesn't have options.");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred while loading state codes: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {

            // Clear and reset the Status Code dropdown
            //comboBox3.Items.Clear();
            //comboBox3.SelectedItem = null;
            //comboBox3.Items.Remove(comboBox3.SelectedIndex);
            comboBox3.ResetText();

            if (comboBox2.SelectedItem != null)
            {
                // Parse the selected statecode value.
                var selectedStateCodeText = comboBox2.SelectedItem.ToString();
                var selectedStateCode = int.Parse(selectedStateCodeText.Split(' ')[0]); // Extract the numeric value.

                //Getting the selected entity
                string selectedEntity = comboBox1.SelectedItem.ToString();

                //Load status codes based on the selected statecode
                LoadStatusCodes(selectedEntity, selectedStateCode);
            }
        }

        private void LoadStatusCodes(string entityLogicalName, int selectedStateCode)
        {
            try
            {
                comboBox3.Items.Clear();
                // Retrieve entity metadata
                var entityMetadataRequest = new RetrieveEntityRequest
                {
                    EntityFilters = EntityFilters.Attributes,
                    LogicalName = entityLogicalName
                };

                RetrieveEntityResponse entityMetadataResponse = (RetrieveEntityResponse)service.Execute(entityMetadataRequest);

                // Find the statuscode attribute in the entity metadata
                var statusCodeAttributeMetadata = entityMetadataResponse.EntityMetadata.Attributes
                    .FirstOrDefault(a => a.LogicalName == "statuscode");

                if (statusCodeAttributeMetadata != null && statusCodeAttributeMetadata is StatusAttributeMetadata statusAttributeMetadata)
                {
                    comboBox3.Items.Clear(); // Clear existing items in the dropdown

                    // Loop through statuscode options and filter by the selected statecode
                    var statusCodeOptions = statusAttributeMetadata.OptionSet.Options;
                    foreach (var option in statusCodeOptions)
                    {
                        if (option is StatusOptionMetadata statusOptionMetadata &&
                            statusOptionMetadata.State == selectedStateCode)
                        {
                            string label = statusOptionMetadata.Label.UserLocalizedLabel.Label;
                            int value = statusOptionMetadata.Value.Value;
                            comboBox3.Items.Add($"{value} {label}");
                        }
                    }

                    // Show a message if no valid options were found
                    if (comboBox3.Items.Count == 0)
                    {
                        MessageBox.Show("No status codes found for the selected state code.", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                else
                {
                    MessageBox.Show("Statuscode field is not available for the selected entity.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred while loading status codes: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        private void UpdateSelectedCount()
        {
            int selectedCount = 0;
            // Loop through all rows and count the selected checkboxes
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                var checkBoxCell = row.Cells["Select"] as DataGridViewCheckBoxCell;
                if (checkBoxCell != null && checkBoxCell.Value != null && (bool)checkBoxCell.Value)
                {
                    selectedCount++;
                }
            }
            // Update the label with the selected count
            label8.Text = $"Selected: {selectedCount} / {textBox1.Text}";
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                if (!(button1.Tag is List<string> successfulRecordIds) || successfulRecordIds.Count == 0)
                {
                    MessageBox.Show("No successful records available for download.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                // Create an Excel package to generate the file 
                using (var package = new ExcelPackage())
                {
                    // Create a worksheet
                    var worksheet = package.Workbook.Worksheets.Add("SuccessfulRecords");

                    // Add headers to the Excel file 
                    var dataTable = (DataTable)dataGridView1.DataSource;
                    for (int col = 0; col < dataTable.Columns.Count; col++)
                    {
                        worksheet.Cells[1, col + 1].Value = dataTable.Columns[col].ColumnName;
                    }

                    // Add the data of successful records to the worksheet
                    int rowIndex = 2;

                    //Get the selected column name from the ComboBox
                    string selectedColumn = comboBox4.SelectedItem.ToString();
                    
                    foreach (var recordId in successfulRecordIds)
                    {
                        var row = dataTable.AsEnumerable().FirstOrDefault(r => r[selectedColumn].ToString() == recordId);
                        if (row != null)
                        {
                            for (int col = 0; col < dataTable.Columns.Count; col++)
                            {
                                worksheet.Cells[rowIndex, col + 1].Value = row[col];
                            }
                            rowIndex++;
                        }
                    }

                    // Ask the users where to save the file
                    using (SaveFileDialog saveFileDialog = new SaveFileDialog())
                    {
                        saveFileDialog.Filter = "Excel Files|*.xlsx";
                        if (saveFileDialog.ShowDialog() == DialogResult.OK)
                        {
                            // Save the Excel file
                            FileInfo fi = new FileInfo(saveFileDialog.FileName);
                            package.SaveAs(fi);
                            MessageBox.Show("Successful records downloaded successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred while downloading the records: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        //Function to close the tool
        private void button2_Click(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                // check if richTextBox1 has any content 
                if (string.IsNullOrEmpty(richTextBox1.Text))
                {
                    MessageBox.Show("There are no error logs to download...", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                // Open a SaveFileDialog to allow the user to choose where to save the file 
                using (SaveFileDialog saveFileDialog = new SaveFileDialog())
                {
                    saveFileDialog.Filter = "Text Files (*.txt)|*.txt"; // Filter for text files
                    saveFileDialog.DefaultExt = "txt"; // Default extension

                    // If the user selects a file and clicks OK
                    if (saveFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        //Get the selected file path
                        string filePath = saveFileDialog.FileName;

                        //Write the contents of richTextBox1 to the selected file 
                        File.WriteAllText(filePath, richTextBox1.Text);

                        //show a message confirming the file has been saved
                        MessageBox.Show("Error log saved successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            catch (Exception ex) 
            {
                // Handle any exceptions that occur during the save process
                MessageBox.Show($"An error occurred while saving the log file: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void toolStripButton3_Click(object sender, EventArgs e)
        {
            CloseTool();
        }



        
        private Image ResizeImage(Image image, int width, int height)
        {
            Bitmap resizedBitmap = new Bitmap(width, height);
            using (Graphics g = Graphics.FromImage(resizedBitmap))
            {
                g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                g.DrawImage(image, 0, 0, width, height);
            }
            return resizedBitmap;
        }

        private void progressBar1_Click(object sender, EventArgs e)
        {

        }
    }
}