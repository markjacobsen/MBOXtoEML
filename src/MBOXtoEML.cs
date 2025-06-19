using System;
using System.IO;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace CFG2.MBOXtoEML;

class MBOXtoEML
{
    public static async Task Main(string[] args)
    {
        Console.WriteLine("MBOX to EML Converter");
        Console.WriteLine("---------------------");

        // --- Argument Parsing ---
        if (args.Length < 2)
        {
            Console.WriteLine("Usage: MboxToEmlConverter <MBOX_FilePath> <Output_Directory>");
            Console.WriteLine("Example: MboxToEmlConverter my_emails.mbox C:\\ExtractedEmails");
            return;
        }

        string mboxFilePath = args[0];
        string outputDirectory = args[1];

        // --- Input Validation ---
        if (!File.Exists(mboxFilePath))
        {
            Console.WriteLine($"Error: MBOX file not found at '{mboxFilePath}'");
            return;
        }

        try
        {
            // Ensure the output directory exists, create if not
            if (!Directory.Exists(outputDirectory))
            {
                Directory.CreateDirectory(outputDirectory);
                Console.WriteLine($"Output directory created: '{outputDirectory}'");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error creating output directory: {ex.Message}");
            return;
        }

        Console.WriteLine($"Processing MBOX file: '{mboxFilePath}'");
        Console.WriteLine($"Saving EML files to: '{outputDirectory}'");
        Console.WriteLine();

        // --- MBOX Processing Logic ---
        await ConvertMboxToEml(mboxFilePath, outputDirectory);

        Console.WriteLine("\nConversion complete. Press any key to exit.");
        Console.ReadKey();
    }

    /// <summary>
    /// Reads an MBOX file line by line, identifies individual email messages,
    /// and saves each message as a separate EML file.
    /// </summary>
    /// <param name="mboxFilePath">The full path to the MBOX file.</param>
    /// <param name="outputDirectory">The directory where EML files will be saved.</param>
    private static async Task ConvertMboxToEml(string mboxFilePath, string outputDirectory)
    {
        int emailCount = 0;
        StringWriter currentEmailContent = null; // Buffer to hold content of the current email

        try
        {
            // Read the MBOX file using a StreamReader
            using (StreamReader reader = new StreamReader(mboxFilePath))
            {
                string line;
                bool inEmail = false; // Flag to indicate if we are currently inside an email message

                while ((line = await reader.ReadLineAsync()) != null)
                {
                    // MBOX files use "From " as a delimiter for new messages.
                    // This line indicates the start of a *new* email.
                    // A standard MBOX delimiter line looks like: "From sender@example.com Tue Jan 01 00:00:00 2020"
                    // We need to be careful not to mistake "From " appearing in the email body as a delimiter.
                    // Such occurrences are typically prefixed with a ">" (e.g., ">From some_user@example.com").
                    if (line.StartsWith("From ") && !line.StartsWith(">From "))
                    {
                        // If we were already in an email, it means the previous email has ended.
                        if (inEmail && currentEmailContent != null)
                        {
                            // Save the completed email content to an EML file
                            await SaveEmlFile(outputDirectory, emailCount, currentEmailContent.ToString());
                        }

                        // Start a new email
                        emailCount++;
                        currentEmailContent = new StringWriter();
                        inEmail = true;

                        // The "From " delimiter line itself is typically NOT part of the EML content.
                        // If your specific MBOX file or a particular EML parser requires it,
                        // you can uncomment the line below to include it.
                        // await currentEmailContent.WriteLineAsync(line);
                    }
                    else if (inEmail && currentEmailContent != null)
                    {
                        // If we are currently in an email, append the line to its content.
                        await currentEmailContent.WriteLineAsync(line);
                    }
                    // If not in an email and not a "From " line, it's likely leading garbage or a malformed file.
                    // We just skip these lines until the first "From " is found.
                }

                // After the loop, save the very last email if one was being processed.
                if (inEmail && currentEmailContent != null)
                {
                    await SaveEmlFile(outputDirectory, emailCount, currentEmailContent.ToString());
                }
            }
            Console.WriteLine($"Successfully extracted {emailCount} emails.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred during MBOX processing: {ex.Message}");
            Console.WriteLine(ex.StackTrace);
        }
    }

    /// <summary>
    /// Saves the provided email content as an EML file in the specified directory.
    /// </summary>
    /// <param name="outputDirectory">The directory to save the file in.</param>
    /// <param name="emailIndex">The index of the email, used for file naming.</param>
    /// <param name="emailContent">The full content of the email.</param>
    private static async Task SaveEmlFile(string outputDirectory, int emailIndex, string emailContent)
    {
        // Generate a filename for the EML file
        // You can enhance this to try and extract Subject or Date from content for better naming.
        string fileName = $"email_{emailIndex:D4}.eml"; // e.g., email_0001.eml, email_0002.eml
        string filePath = Path.Combine(outputDirectory, fileName);

        try
        {
            await File.WriteAllTextAsync(filePath, emailContent.Trim()); // Trim to remove any trailing newlines
            Console.WriteLine($"Saved: {fileName}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error saving {fileName}: {ex.Message}");
        }
    }
}
