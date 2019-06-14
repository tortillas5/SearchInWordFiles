// <copyright file="Program.cs" company="Tortillas-Inc">
// Copy me, no rights reserved.
// </copyright>

namespace SearchInWordFiles
{
    using System;
    using System.Collections.Generic;
    using System.IO;

    /// <summary>
    /// Class containing the main methods used to search among word files.
    /// </summary>
    internal class Program
    {
        /// <summary>
        /// Method that allows to choose what type of research is being used.
        /// </summary>
        public static void SelectFonction()
        {
            // Location of the .exe file.
            string path = System.Reflection.Assembly.GetExecutingAssembly().Location;

            // Folder of the .exe file.
            string directory = System.IO.Path.GetDirectoryName(path);
            DirectoryInfo directoryInfo = new DirectoryInfo(directory);

            Console.WriteLine("Types de recherche :");
            Console.WriteLine("1 : Rechercher une région par correspondance exacte.");
            Console.WriteLine("2 : Rechercher une région par correspondance incomplète.");
            Console.WriteLine("3 : Rechercher partout un mot / une phrase.");

            string selectedFunction = ReadWriteConsoleManagement.GetSaisieUtilisateur(
                startMessage: "Choisissez le type de recherche :",
                errorMessage: "Cette recherche n'existe pas.",
                acceptedValues: new List<string>() { "1", "2", "3" });

            Console.Clear();

            switch (selectedFunction)
            {
                case "1":
                    string regionComplete = ReadWriteConsoleManagement.GetSaisieUtilisateur("Rentrez la région complète (en respectant la case) à chercher parmi les documents word du dossier :", "Veuillez rentrer la région (en respectant la case).");
                    WordSearch.RecursiveSearchRegionInWordFiles(directoryInfo, regionComplete);
                    break;

                case "2":
                    string regionPartielle = ReadWriteConsoleManagement.GetSaisieUtilisateur("Rentrez un mot appartenant à la région à chercher parmi les documents word du dossier :", "Veuillez rentrer un mot appartenant à la région.");
                    WordSearch.RecursiveSearchRegionInWordFilesWithPartialWord(directoryInfo, regionPartielle);
                    break;

                case "3":
                    string motPhrase = ReadWriteConsoleManagement.GetSaisieUtilisateur("Rentrez un mot / une phrase à chercher parmi les documents word du dossier :", "Veuillez rentrer un mot / une phrase.");
                    WordSearch.RecursiveSearchWordInFiles(directoryInfo, motPhrase);
                    break;
            }
        }

        /// <summary>
        /// Main method, launched first.
        /// </summary>
        /// <param name="args">Arguments of the program (not used. Yet ?).</param>
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
                Console.WriteLine(ExceptionsHandler.ConcatException(ex));
                Console.ReadLine();
            }
        }
    }
}