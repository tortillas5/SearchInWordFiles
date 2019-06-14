// <copyright file="WordSearch.cs" company="Tortillas-Inc">
// Copy me, no rights reserved.
// </copyright>

namespace SearchInWordFiles
{
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Linq;
    using Aspose.Words;

    /// <summary>
    /// Class containing the main methods used to search among word files.
    /// </summary>
    public static class WordSearch
    {
        /// <summary>
        /// Search all the word files (.doc / .docx) in the directoryInfo folder to see if a region contains the word to find.
        /// Also search among the subfolders of the current file.
        /// </summary>
        /// <param name="directoryInfo">Folder in which we will search on each file .doc / .docx.</param>
        /// <param name="searchedWord">Word to find.</param>
        public static void RecursiveSearchRegionInWordFiles(DirectoryInfo directoryInfo, string searchedWord)
        {
            // Load option of a document.
            LoadOptions loadOptions = new LoadOptions
            {
                LoadFormat = LoadFormat.Doc
            };

            // For the files of the folder.
            foreach (FileInfo fileInfo in directoryInfo.GetFiles())
            {
                if (fileInfo.Extension == ".doc" || fileInfo.Extension == ".docx")
                {
                    Console.WriteLine("Recherche dans le document " + fileInfo.Name + '.');

                    using (FileStream fileStream = System.IO.File.OpenRead(fileInfo.FullName))
                    {
                        Document document = new Document(fileStream, loadOptions);
                        List<string> regionsDocument = document
                            .MailMerge
                            .GetFieldNames()
                            .Where(fn => fn.StartsWith("TableStart"))
                            .Select(reg => reg.Substring(reg.IndexOf(":") + 1))
                            .ToList();

                        if (regionsDocument.Contains(searchedWord))
                        {
                            Console.WriteLine("La région " + searchedWord + " a été trouvée dans le fichier " + fileInfo.Name + ".");
                            Console.WriteLine("Voulez vous continuer la recherche (Y), ou l'arrêter ? (N)");

                            string response = ReadWriteConsoleManagement.GetSaisieUtilisateur("Y / N ?", "Veuillez rentrer Y (Yes) / N (No).", new List<string>() { "Y", "N" });

                            if (response == "N")
                            {
                                break;
                            }
                            else
                            {
                                continue;
                            }
                        }
                    }
                }
            }

            // For the subfolders.
            foreach (DirectoryInfo di in directoryInfo.GetDirectories())
            {
                Console.WriteLine("Recherche dans le dossier " + di.FullName);

                WordSearch.RecursiveSearchRegionInWordFiles(di, searchedWord);
            }
        }

        /// <summary>
        /// Search all the word files (.doc / .docx) in the directoryInfo folder to see if a region contains the partial word to find.
        /// Also search among the subfolders of the current file.
        /// </summary>
        /// <param name="directoryInfo">Folder in which we will search on each file .doc / .docx.</param>
        /// <param name="searchedPartialWord">Partial word to find.</param>
        public static void RecursiveSearchRegionInWordFilesWithPartialWord(DirectoryInfo directoryInfo, string searchedPartialWord)
        {
            // Load option of a document.
            LoadOptions loadOptions = new LoadOptions
            {
                LoadFormat = LoadFormat.Doc
            };

            // For the files of the folder.
            foreach (FileInfo fileInfo in directoryInfo.GetFiles())
            {
                if (fileInfo.Extension == ".doc" || fileInfo.Extension == ".docx")
                {
                    Console.WriteLine("Recherche dans le document " + fileInfo.Name + '.');

                    using (FileStream fileStream = System.IO.File.OpenRead(fileInfo.FullName))
                    {
                        Document document = new Document(fileStream, loadOptions);
                        List<string> regionsDocument = document
                            .MailMerge
                            .GetFieldNames()
                            .Where(fn => fn.StartsWith("TableStart"))
                            .Select(reg => reg.Substring(reg.IndexOf(":") + 1))
                            .ToList();

                        if (regionsDocument.Any(rd => rd.Contains(searchedPartialWord)))
                        {
                            Console.WriteLine("Une région contenant le mot " + searchedPartialWord + " a été trouvée dans le fichier " + fileInfo.Name + ".");
                            Console.WriteLine("Voulez vous continuer la recherche (Y), ou l'arrêter ? (N)");

                            string response = ReadWriteConsoleManagement.GetSaisieUtilisateur("Y / N ?", "Veuillez rentrer Y (Yes) / N (No).", new List<string>() { "Y", "N" });

                            if (response == "N")
                            {
                                break;
                            }
                            else
                            {
                                continue;
                            }
                        }
                    }
                }
            }

            // For the subfolders.
            foreach (DirectoryInfo di in directoryInfo.GetDirectories())
            {
                Console.WriteLine("Recherche dans le dossier " + di.FullName);

                WordSearch.RecursiveSearchRegionInWordFiles(di, searchedPartialWord);
            }
        }

        /// <summary>
        /// Search all the word files (.doc / .docx) in the directoryInfo folder to see if any contains the word / phrase to find.
        /// Also search among the subfolders of the current file.
        /// </summary>
        /// <param name="directoryInfo">Folder in which we will search on each file .doc / .docx.</param>
        /// <param name="searchedPartialWord">Word / phrase to find.</param>
        public static void RecursiveSearchWordInFiles(DirectoryInfo directoryInfo, string searchedPartialWord)
        {
            // Load option of a document.
            LoadOptions loadOptions = new LoadOptions
            {
                LoadFormat = LoadFormat.Doc
            };

            // For the files of the folder.
            foreach (FileInfo fileInfo in directoryInfo.GetFiles())
            {
                if (fileInfo.Extension == ".doc" || fileInfo.Extension == ".docx")
                {
                    Console.WriteLine("Recherche dans le document " + fileInfo.Name + '.');

                    using (FileStream fileStream = System.IO.File.OpenRead(fileInfo.FullName))
                    {
                        Document document = new Document(fileStream, loadOptions);
                        string textDocument = document
                            .GetText();

                        if (textDocument.Contains(searchedPartialWord))
                        {
                            Console.WriteLine("Un fichier contenant le mot / la phrase " + searchedPartialWord + " a été trouvé : " + fileInfo.Name + ".");
                            Console.WriteLine("Voulez vous continuer la recherche (Y), ou l'arrêter ? (N)");

                            string response = ReadWriteConsoleManagement.GetSaisieUtilisateur("Y / N ?", "Veuillez rentrer Y (Yes) / N (No).", new List<string>() { "Y", "N" });

                            if (response == "N")
                            {
                                break;
                            }
                            else
                            {
                                continue;
                            }
                        }
                    }
                }
            }

            // For the subfolders.
            foreach (DirectoryInfo di in directoryInfo.GetDirectories())
            {
                Console.WriteLine("Recherche dans le dossier " + di.FullName);

                WordSearch.RecursiveSearchRegionInWordFiles(di, searchedPartialWord);
            }
        }
    }
}