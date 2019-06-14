// <copyright file="Program.cs" company="Tortillas-Inc">
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
    /// Classe contenant les principales méthodes utilisée pour chercher parmi des fichiers word.
    /// </summary>
    internal class Program
    {
        /// <summary>
        /// Cherche parmi tous les fichiers word (doc / docx) du dossier directoryInfo, pour voir si une région contient le mot à trouver.
        /// Cherche aussi parmi les sous dossiers du dossier en cours.
        /// </summary>
        /// <param name="directoryInfo">Dossier dans lequel on va chercher sur chaque fichier.</param>
        /// <param name="searchedWord">Mot à trouver.</param>
        public static void RecursiveSearchRegionInWordFiles(DirectoryInfo directoryInfo, string searchedWord)
        {
            // Option de chargement d'un document
            LoadOptions loadOptions = new LoadOptions
            {
                LoadFormat = LoadFormat.Doc
            };

            // Pour les fichiers du dossier.
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

                            string response = GestionReadWriteConsole.GetSaisieUtilisateur("Y / N ?", "Veuillez rentrer Y (Yes) / N (No).", new List<string>() { "Y", "N" });

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

            // Pour les sous dossiers du dossier.
            foreach (DirectoryInfo di in directoryInfo.GetDirectories())
            {
                Console.WriteLine("Recherche dans le dossier " + di.FullName);

                Program.RecursiveSearchRegionInWordFiles(di, searchedWord);
            }
        }

        /// <summary>
        /// Cherche parmi tous les fichiers word (doc / docx) du dossier directoryInfo, pour voir si une région contient le mot partiel à trouver.
        /// Cherche aussi parmi les sous dossiers du dossier en cours.
        /// </summary>
        /// <param name="directoryInfo">Dossier dans lequel on va chercher sur chaque fichier.</param>
        /// <param name="searchedPartialWord">Mot partiel à trouver.</param>
        public static void RecursiveSearchRegionInWordFilesWithPartialWord(DirectoryInfo directoryInfo, string searchedPartialWord)
        {
            // Option de chargement d'un document
            LoadOptions loadOptions = new LoadOptions
            {
                LoadFormat = LoadFormat.Doc
            };

            // Pour les fichiers du dossier.
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

                            string response = GestionReadWriteConsole.GetSaisieUtilisateur("Y / N ?", "Veuillez rentrer Y (Yes) / N (No).", new List<string>() { "Y", "N" });

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

            // Pour les sous dossiers du dossier.
            foreach (DirectoryInfo di in directoryInfo.GetDirectories())
            {
                Console.WriteLine("Recherche dans le dossier " + di.FullName);

                Program.RecursiveSearchRegionInWordFiles(di, searchedPartialWord);
            }
        }

        /// <summary>
        /// Cherche parmi tous les fichiers word (doc / docx) du dossier directoryInfo, pour voir si l'un d'entre eux contient le mot / la phrase à trouver.
        /// Cherche aussi parmi les sous dossiers du dossier en cours.
        /// </summary>
        /// <param name="directoryInfo">Dossier dans lequel on va chercher sur chaque fichier.</param>
        /// <param name="searchedPartialWord">Mot /phrase à trouver.</param>
        public static void RecursiveSearchWordInFiles(DirectoryInfo directoryInfo, string searchedPartialWord)
        {
            // Option de chargement d'un document
            LoadOptions loadOptions = new LoadOptions
            {
                LoadFormat = LoadFormat.Doc
            };

            // Pour les fichiers du dossier.
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

                            string response = GestionReadWriteConsole.GetSaisieUtilisateur("Y / N ?", "Veuillez rentrer Y (Yes) / N (No).", new List<string>() { "Y", "N" });

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

            // Pour les sous dossiers du dossier.
            foreach (DirectoryInfo di in directoryInfo.GetDirectories())
            {
                Console.WriteLine("Recherche dans le dossier " + di.FullName);

                Program.RecursiveSearchRegionInWordFiles(di, searchedPartialWord);
            }
        }

        /// <summary>
        /// Méthode qui permet de choisir quel type de recherche on va utiliser.
        /// </summary>
        public static void SelectFonction()
        {
            // Emplacement de l'exécutable
            string path = System.Reflection.Assembly.GetExecutingAssembly().Location;

            // Répertoire de l'exécutable
            string directory = System.IO.Path.GetDirectoryName(path);
            DirectoryInfo directoryInfo = new DirectoryInfo(directory);

            Console.WriteLine("Types de recherche :");
            Console.WriteLine("1 : Rechercher une région par correspondance exacte.");
            Console.WriteLine("2 : Rechercher une région par correspondance incomplète.");
            Console.WriteLine("3 : Rechercher partout un mot / une phrase.");

            string selectedFunction = GestionReadWriteConsole.GetSaisieUtilisateur(
                startMessage: "Choisissez le type de recherche :",
                errorMessage: "Cette recherche n'existe pas.",
                acceptedValues: new List<string>() { "1", "2", "3" });

            Console.Clear();

            switch (selectedFunction)
            {
                case "1":
                    string regionComplete = GestionReadWriteConsole.GetSaisieUtilisateur("Rentrez la région complète (en respectant la case) à chercher parmi les documents word du dossier :", "Veuillez rentrer la région (en respectant la case).");
                    Program.RecursiveSearchRegionInWordFiles(directoryInfo, regionComplete);
                    break;

                case "2":
                    string regionPartielle = GestionReadWriteConsole.GetSaisieUtilisateur("Rentrez un mot appartenant à la région à chercher parmi les documents word du dossier :", "Veuillez rentrer un mot appartenant à la région.");
                    Program.RecursiveSearchRegionInWordFilesWithPartialWord(directoryInfo, regionPartielle);
                    break;

                case "3":
                    string motPhrase = GestionReadWriteConsole.GetSaisieUtilisateur("Rentrez un mot / une phrase à chercher parmi les documents word du dossier :", "Veuillez rentrer un mot / une phrase.");
                    Program.RecursiveSearchWordInFiles(directoryInfo, motPhrase);
                    break;
            }
        }

        /// <summary>
        /// Méthode principale, lancée en première.
        /// </summary>
        /// <param name="args">Arguments du programme.</param>
        private static void Main(string[] args)
        {
            try
            {
                Program.SelectFonction();

                Console.WriteLine("Recherche terminée.");
                Console.ReadLine();
            }
            catch (Exception ex)
            {
                Console.WriteLine(GererExceptions.ConcatException(ex));
                Console.ReadLine();
            }
        }
    }
}