using Aspose.Words;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using Microsoft.Azure.Storage;
using Microsoft.Azure.Storage.Blob;
using System.Data.SqlClient;
using System.Text.RegularExpressions;

class Program
{
    static void Main()
    {
        int totalFileUploded = 0;
        int successfullyUploaded = 0;
        int failedToUpload = 0;
        int notAResume = 0;
        // Folder location
        string folderPath = @"C:\Users\POOJA-1\Desktop\myResume";
        // Blob storage location
        string storageConnectionString = "DefaultEndpointsProtocol=https;AccountName=strecruitmentportal;AccountKey=OwBjNn+cByJwj/ht8rpd5At7takptlPv2OvW8fGYJrjVnhS/PBFYJ0gvOyn77a/bgSZYjSSf5ixN+AStBzhY7g==;EndpointSuffix=core.windows.net;";
        // Az.Sql database location
        string sqlConnectionString = "Server=tcp:sql-prod-recruitmentportal-centralindia-01.database.windows.net,1433;Initial Catalog=sqldb-prod-recruitmentportal-centralindia-01;Persist Security Info=False;User ID=wbrpprodadmin;Password=!RecruitmentPortalProd1;MultipleActiveResultSets=False;Encrypt=True;TrustServerCertificate=False;Connection Timeout=30;";






        FileSystemWatcher watcher = new FileSystemWatcher(folderPath);
        watcher.Created += (sender, e) =>
        {
            string filePath = e.FullPath;
            totalFileUploded++;  //Count the total uploads

            string pattern = @"\\([^\\]+)$";
            MatchCollection matches = Regex.Matches(filePath, pattern);

            string fileName = matches[0].Groups[1].Value;

            // Parse the file and extract email address
            string emailAddress = ParseEmailAddress(filePath);

            if (emailAddress != null)
            {

                // Store email and blob URL in SQL database
                using (SqlConnection connection = new SqlConnection(sqlConnectionString))
                {
                    connection.Open();

                    // Check if candidate with matching email exists
                    string checkQuery = "SELECT COUNT(*) FROM Candidate WHERE EmailId LIKE '%' + @Email + '%'";
                    using (SqlCommand checkCommand = new SqlCommand(checkQuery, connection))
                    {
                        checkCommand.Parameters.AddWithValue("@Email", emailAddress);
                        int existingCount = (int)checkCommand.ExecuteScalar();

                        if (existingCount > 0)
                        {
                            // Candidate with matching email exists, update their information
                            // Upload the resume to Azure Blob Storage
                            CloudStorageAccount storageAccount = CloudStorageAccount.Parse(storageConnectionString);
                            CloudBlobClient blobClient = storageAccount.CreateCloudBlobClient();
                            //Resume are stored in resumes container in the blob
                            CloudBlobContainer container = blobClient.GetContainerReference("candidateresume");
                            container.CreateIfNotExists();

                            CloudBlockBlob blob = container.GetBlockBlobReference(System.IO.Path.GetFileName(filePath));
                            using (FileStream fs = File.OpenRead(filePath))
                            {
                                blob.UploadFromStream(fs);
                            }
                            string updateQuery = "UPDATE Candidate SET Resume = @Url WHERE EmailId LIKE '%' + @Email + '%'";
                            using (SqlCommand updateCommand = new SqlCommand(updateQuery, connection))
                            {
                                updateCommand.Parameters.AddWithValue("@Email", emailAddress);
                                updateCommand.Parameters.AddWithValue("@Url", blob.StorageUri.PrimaryUri.ToString());
                                updateCommand.ExecuteNonQuery();
                                successfullyUploaded++;
                            }


                            Console.WriteLine(totalFileUploded + "," + fileName + ",Uploaded," + emailAddress + ",Found");
                        }
                        else
                        {
                            failedToUpload++;
                            // Candidate with email doesn't exist, insert new record
                            Console.WriteLine(totalFileUploded + "," + fileName + ",Not Uploaded," + emailAddress + ",Not Found");
                        }
                    }
                }
            }
            else
            {
                Console.WriteLine(totalFileUploded + "," + fileName + ",Not Uploaded,NA,NA");
                notAResume++;
            }
        };

        watcher.EnableRaisingEvents = true;

        Console.WriteLine("Press 'q' to quit.");
        while (Console.Read() != 'q') ;
        // To print the final report 
        Console.WriteLine("\nTotal File Uploaded: " + totalFileUploded);
        Console.WriteLine("Ignored File which is not recognised as resume count: " + notAResume);
        Console.WriteLine("Successfully Uploaded Resume count  : " + successfullyUploaded);
        Console.WriteLine("Failed to upload due to email not available in database : " + failedToUpload);

    }


    // Extract the text from the file
    static string ParseEmailAddress(string filePath)
    {
        string emailAddress = null;

        // Pdf file parsing using iTextSharper
        if (System.IO.Path.GetExtension(filePath).Equals(".pdf", StringComparison.OrdinalIgnoreCase))
        {
            try
            {
                using (PdfReader pdfReader = new PdfReader(filePath))
                {
                    for (int page = 1; page <= pdfReader.NumberOfPages; page++)
                    {
                        string text = PdfTextExtractor.GetTextFromPage(pdfReader, page);

                        //parse the email using the Regex
                        emailAddress = ExtractEmailAddress(text);
                        if (!string.IsNullOrEmpty(emailAddress))
                            break;
                    }
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }

        // Word file parsing using Aspose.Words
        else if (System.IO.Path.GetExtension(filePath).Equals(".docx", StringComparison.OrdinalIgnoreCase) ||
            System.IO.Path.GetExtension(filePath).Equals(".doc", StringComparison.OrdinalIgnoreCase))
        {
            try
            {
                Document doc = new Document(filePath);
                string text = doc.GetText();

                //parse the email using the Regex
                emailAddress = ExtractEmailAddress(text);
            }
            catch
            {
                Console.WriteLine("error at catch block");
            }
        }

        return emailAddress;
    }

    //Extract email address
    static string ExtractEmailAddress(string text)
    {
        // Use regular expressions to find and extract email addresses
        string pattern = @"\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b";
        Match match = Regex.Match(text, pattern);
        return match.Success ? match.Value : null;
    }
}
