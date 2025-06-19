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
    /// It also attempts to extract the Subject and Date headers for naming the EML files.
    /// </summary>
    /// <param name="mboxFilePath">The full path to the MBOX file.</param>
    /// <param name="outputDirectory">The directory where EML files will be saved.</param>
    private static async Task ConvertMboxToEml(string mboxFilePath, string outputDirectory)
    {
        int emailCount = 0;
        StringWriter currentEmailContent = null; // Buffer to hold content of the current email
        string currentSubject = "No Subject";   // Stores the subject of the current email
        string currentSentDate = "No Date";     // Stores the sent date of the current email

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
                    // We check for lines starting with "From " but not ">From " (which would be escaped in content).
                    if (line.StartsWith("From ") && !line.StartsWith(">From "))
                    {
                        // If we were already in an email, it means the previous email has ended.
                        if (inEmail && currentEmailContent != null)
                        {
                            // Save the completed email content to an EML file, using extracted subject and date
                            await SaveEmlFile(outputDirectory, emailCount, currentEmailContent.ToString(), currentSubject, currentSentDate);
                        }

                        // Start a new email
                        emailCount++;
                        currentEmailContent = new StringWriter();
                        inEmail = true;
                        currentSubject = "No Subject";   // Reset subject for the new email
                        currentSentDate = "No Date";     // Reset date for the new email

                        // The "From " delimiter line itself is typically NOT part of the EML content.
                        // It simply marks the beginning of a new message.
                        continue; // Skip adding this delimiter line to the current email content
                    }

                    // If we are currently processing an email, add the line to its content.
                    if (inEmail && currentEmailContent != null)
                    {
                        await currentEmailContent.WriteLineAsync(line);

                        // Attempt to extract Subject and Date from headers.
                        // This simple approach only captures the first occurrence and doesn't handle multi-line headers.
                        if (currentSubject == "No Subject" && line.StartsWith("Subject:", StringComparison.OrdinalIgnoreCase))
                        {
                            currentSubject = SanitizeFilename(line.Substring("Subject:".Length).Trim());
                        }
                        else if (currentSentDate == "No Date" && line.StartsWith("Date:", StringComparison.OrdinalIgnoreCase))
                        {
                            string dateStr = line.Substring("Date:".Length).Trim();
                            // Try to parse the date to a sortable format (YYYY-MM-DD_HHMMSS)
                            if (DateTimeOffset.TryParse(dateStr, out DateTimeOffset parsedDate))
                            {
                                currentSentDate = parsedDate.ToString("yyyy-MM-dd_HHmmss");
                            }
                            else
                            {
                                // If parsing fails, use a sanitized version of the raw date string
                                currentSentDate = SanitizeFilename(dateStr);
                            }
                        }
                    }
                    // If not in an email and not a "From " line, it's likely leading garbage or a malformed file.
                    // We just skip these lines until the first "From " is found.
                }

                // After the loop, save the very last email if one was being processed.
                if (inEmail && currentEmailContent != null)
                {
                    await SaveEmlFile(outputDirectory, emailCount, currentEmailContent.ToString(), currentSubject, currentSentDate);
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
    /// The filename will include the extracted subject and sent date for better organization.
    /// </summary>
    /// <param name="outputDirectory">The directory to save the file in.</param>
    /// <param name="emailIndex">The index of the email, used as a unique identifier in the filename.</param>
    /// <param name="emailContent">The full content of the email.</param>
    /// <param name="subject">The subject of the email, used for filename.</param>
    /// <param name="sentDate">The sent date of the email, formatted for filename.</param>
    private static async Task SaveEmlFile(string outputDirectory, int emailIndex, string emailContent, string subject, string sentDate)
    {
        // Ensure subject and sentDate are not empty or too long for filenames
        string safeSubject = string.IsNullOrWhiteSpace(subject) || subject == "No Subject" ? "NoSubject" : subject;
        string safeDate = string.IsNullOrWhiteSpace(sentDate) || sentDate == "No Date"
                            ? DateTime.Now.ToString("yyyy-MM-dd_HHmmss") // Default to current time if no date found
                            : sentDate;

        // Truncate subject if it's too long for a practical filename (e.g., limit to 50 characters)
        if (safeSubject.Length > 50)
        {
            safeSubject = safeSubject.Substring(0, 50);
        }

        // Construct the filename using date, sanitized subject, and index
        string fileName = $"{safeDate}_{safeSubject}_{emailIndex:D4}.eml"; // e.g., 20231027_153045_MyEmailSubject_0001.eml
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

    /// <summary>
    /// Sanitizes a string to be safe for use as a filename or path component.
    /// Replaces invalid characters with underscores and also replaces spaces with underscores.
    /// </summary>
    /// <param name="input">The string to sanitize.</param>
    /// <returns>A sanitized string suitable for a filename.</returns>
    private static string SanitizeFilename(string input)
    {
        // Get invalid characters for filenames and paths
        char[] invalidChars = Path.GetInvalidFileNameChars();
        char[] invalidPathChars = Path.GetInvalidPathChars();

        // Combine them and escape for regex
        string allInvalidChars = new string(invalidChars) + new string(invalidPathChars);
        string invalidCharsPattern = $"[{Regex.Escape(allInvalidChars)}]";

        // Replace invalid characters with an underscore, and also replace spaces with underscores.
        // Multiple underscores are condensed into a single underscore for cleaner names.
        string sanitized = Regex.Replace(input, invalidCharsPattern, "_");
        sanitized = Regex.Replace(sanitized, @"\s+", "_"); // Replace any whitespace with single underscore
        sanitized = Regex.Replace(sanitized, "_+", "_"); // Replace multiple underscores with single underscore
        sanitized = sanitized.Trim('_'); // Remove leading/trailing underscores

        return sanitized;
    }
}
