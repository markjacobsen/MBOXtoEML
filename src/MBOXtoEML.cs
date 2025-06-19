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

        Console.WriteLine("\nConversion complete.");
    }

    /// <summary>
    /// A helper class to hold the raw lines of an email segment, its extracted metadata,
    /// and a flag to track if it was successfully extracted as an EML.
    /// </summary>
    private class EmailSegment
    {
        public List<string> RawLines { get; set; } = new List<string>();
        public string Subject { get; set; } = "No Subject";
        public string SentDate { get; set; } = "No Date";
        public bool SuccessfullyExtracted { get; set; } = false;
        public string FromDelimiterLine { get; set; } // Stores the original "From " delimiter line
    }

    /// <summary>
    /// Reads an MBOX file line by line, identifies individual email messages,
    /// extracts their subject and date, saves each message as a separate EML file,
    /// and finally rewrites the MBOX file to remove the successfully extracted emails.
    /// </summary>
    /// <param name="mboxFilePath">The full path to the MBOX file.</param>
    /// <param name="outputDirectory">The directory where EML files will be saved.</param>
    private static async Task ConvertMboxToEml(string mboxFilePath, string outputDirectory)
    {
        List<EmailSegment> emailsInMbox = new List<EmailSegment>();
        EmailSegment currentSegment = null;
        int totalEmailsIdentified = 0;

        // --- Phase 1: Read MBOX and Segment Emails into memory ---
        try
        {
            Console.WriteLine("Phase 1/3: Reading MBOX file and segmenting emails...");
            using (StreamReader reader = new StreamReader(mboxFilePath))
            {
                string line;
                while ((line = await reader.ReadLineAsync()) != null)
                {
                    // MBOX files use "From " as a delimiter for new messages.
                    // We check for lines starting with "From " but not ">From " (which would be escaped in content).
                    if (line.StartsWith("From ") && !line.StartsWith(">From "))
                    {
                        // If a previous segment was being processed, add it to the list
                        if (currentSegment != null)
                        {
                            emailsInMbox.Add(currentSegment);
                        }
                        totalEmailsIdentified++;
                        currentSegment = new EmailSegment();
                        currentSegment.FromDelimiterLine = line; // Store the original "From " delimiter line
                    }
                    else if (currentSegment != null)
                    {
                        // Add line to current email segment's raw lines
                        currentSegment.RawLines.Add(line);

                        // Attempt to extract Subject and Date from headers if not already found
                        // This simple approach only captures the first occurrence and doesn't handle multi-line headers.
                        if (currentSegment.Subject == "No Subject" && line.StartsWith("Subject:", StringComparison.OrdinalIgnoreCase))
                        {
                            currentSegment.Subject = SanitizeFilename(line.Substring("Subject:".Length).Trim());
                        }
                        else if (currentSegment.SentDate == "No Date" && line.StartsWith("Date:", StringComparison.OrdinalIgnoreCase))
                        {
                            string dateStr = line.Substring("Date:".Length).Trim();
                            // Try to parse the date to a sortable format (YYYY-MM-DD_HHMMSS)
                            if (DateTimeOffset.TryParse(dateStr, out DateTimeOffset parsedDate))
                            {
                                currentSegment.SentDate = parsedDate.ToString("yyyy-MM-dd_HHmmss");
                            }
                            else
                            {
                                // If parsing fails, use a sanitized version of the raw date string
                                currentSegment.SentDate = SanitizeFilename(dateStr);
                            }
                        }
                    }
                }
                // Add the very last segment if it exists after the loop
                if (currentSegment != null)
                {
                    emailsInMbox.Add(currentSegment);
                }
            }
            Console.WriteLine($"Identified {totalEmailsIdentified} email segments in the MBOX file.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred during MBOX segmentation: {ex.Message}");
            Console.WriteLine(ex.StackTrace);
            return; // Cannot proceed without proper segmentation
        }

        // --- Phase 2: Save EML files and mark successful extractions ---
        Console.WriteLine("Phase 2/3: Saving emails as EML files...");
        int successfullyExtractedCount = 0;
        for (int i = 0; i < emailsInMbox.Count; i++)
        {
            var email = emailsInMbox[i];
            // Reconstruct the email content from its raw lines for saving as EML
            using (StringWriter emailContentWriter = new StringWriter())
            {
                foreach (var line in email.RawLines)
                {
                    await emailContentWriter.WriteLineAsync(line);
                }
                // SaveEmlFile returns true on success, false on failure
                email.SuccessfullyExtracted = await SaveEmlFile(outputDirectory, i + 1, emailContentWriter.ToString(), email.Subject, email.SentDate);
                if (email.SuccessfullyExtracted)
                {
                    successfullyExtractedCount++;
                }
            }
        }
        Console.WriteLine($"Successfully extracted {successfullyExtractedCount} emails to EML files.");


        // --- Phase 3: Rewrite MBOX file to remove extracted emails ---
        try
        {
            Console.WriteLine("Phase 3/3: Rewriting MBOX file to remove successfully extracted emails...");
            int keptEmailsCount = 0;
            // Open the MBOX file for writing, overwriting its content (false means truncate/overwrite)
            using (StreamWriter writer = new StreamWriter(mboxFilePath, false))
            {
                foreach (var email in emailsInMbox)
                {
                    // If the email was NOT successfully extracted, write it back to the MBOX
                    if (!email.SuccessfullyExtracted)
                    {
                        // Re-add the "From " delimiter line that was stripped during initial parsing
                        if (!string.IsNullOrEmpty(email.FromDelimiterLine))
                        {
                            await writer.WriteLineAsync(email.FromDelimiterLine);
                        }
                        foreach (var line in email.RawLines)
                        {
                            await writer.WriteLineAsync(line);
                        }
                        keptEmailsCount++;
                    }
                }
            }
            Console.WriteLine($"MBOX file rewritten. {keptEmailsCount} emails remain (those that failed to extract).");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error rewriting MBOX file: {ex.Message}");
            Console.WriteLine("The original MBOX file might be partially modified or corrupted. Please check it manually.");
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
    /// <returns>True if the EML file was saved successfully, false otherwise.</returns>
    private static async Task<bool> SaveEmlFile(string outputDirectory, int emailIndex, string emailContent, string subject, string sentDate)
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
            return true; // Indicate success
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error saving {fileName}: {ex.Message}");
            return false; // Indicate failure
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
