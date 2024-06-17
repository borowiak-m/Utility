public void Main()
        {
            // Get the initial folder path from the SSIS variable
            string initialFolderPath = Dts.Variables["User::FolderPath"].Value.ToString();
            MessageBox.Show("Initial folder path:" + initialFolderPath);

            // Get the financial year from the SSIS variable
            string financialYear = Dts.Variables["User::Reported_year"].Value.ToString();
            MessageBox.Show("Financial year:" + financialYear);

            string outputFilesFolderPath = Path.Combine(initialFolderPath, "FY" + financialYear, "Output Files", "FD");
            string masterFilesFolderPath = Path.Combine(initialFolderPath, "FY" + financialYear, "Master Files");
            MessageBox.Show("outputFilesFolderPath =" + outputFilesFolderPath);
            MessageBox.Show(" masterFilesFolderPathh =" + masterFilesFolderPath);

            // Check and create the "Output Files\FD" folder structure
            CreateFolderStructure(outputFilesFolderPath, "Output Files\\FD");

            // Check and create the "Master Files" folder structure
            CreateFolderStructure(masterFilesFolderPath, "Master Files");

            Dts.TaskResult = (int)ScriptResults.Success;
        }

        private void CreateFolderStructure(string folderPath, string folderName)
        {
            bool fireAgain = true;

            if (Directory.Exists(folderPath))
            {
                // Folder structure exists, log a message
                Dts.Events.FireInformation(0, "FolderCheck", $"Folder structure exists: {folderPath}", string.Empty, 0, ref fireAgain);
            }
            else
            {
                // Folder structure does not exist, create it
                try
                {
                    Directory.CreateDirectory(folderPath);
                    Dts.Events.FireInformation(0, "FolderCheck", $"Folder structure created: {folderPath}", string.Empty, 0, ref fireAgain);
                }
                catch (Exception ex)
                {
                    // Handle any exceptions that occurred during folder creation
                    Dts.Events.FireError(0, "FolderCheck", ex.Message, string.Empty, 0);
                }
            }
        }
