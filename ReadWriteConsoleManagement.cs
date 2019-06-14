// <copyright file="ReadWriteConsoleManagement.cs" company="Tortillas-Inc">
// Copy me, no rights reserved.
// </copyright>

namespace SearchInWordFiles
{
    using System;
    using System.Collections.Generic;
    using System.Linq;

    /// <summary>
    /// Class containing the methods handling reading / writing into console, making it easier to interact with the user.
    /// </summary>
    public static class ReadWriteConsoleManagement
    {
        /// <summary>
        /// Ask the user to input text. Check is the text is correct. Check if the text is null or empty, or white spaces.
        /// </summary>
        /// <param name="startMessage">Start message, telling the user what to enter as a parameter.</param>
        /// <param name="errorMessage">Error message if the user enter a wrong value into the console.</param>
        /// <returns>Text entered by the user. Has been checked.</returns>
        public static string GetSaisieUtilisateur(string startMessage, string errorMessage)
        {
            string searchedWord = string.Empty;

            do
            {
                Console.WriteLine(startMessage);
                searchedWord = Console.ReadLine();

                if (string.IsNullOrWhiteSpace(searchedWord))
                {
                    Console.WriteLine(errorMessage);
                    searchedWord = string.Empty;
                }
            }
            while (string.IsNullOrWhiteSpace(searchedWord));

            return searchedWord;
        }

        /// <summary>
        /// Ask the user to input text. Check is the text is correct. Check that the entered text matches a list of predefined values.
        /// </summary>
        /// <param name="startMessage">Start message, telling the user what to enter as a parameter.</param>
        /// <param name="errorMessage">Error message if the user enter a wrong value into the console.</param>
        /// <param name="acceptedValues">List of accepted values in the user input.</param>
        /// <returns>Text entered by the user. Has been checked.</returns>
        public static string GetSaisieUtilisateur(string startMessage, string errorMessage, List<string> acceptedValues)
        {
            if (!acceptedValues.Any())
            {
                throw new Exception("La liste des valeurs acceptées est vide.");
            }

            string searchedWord = string.Empty;

            do
            {
                Console.WriteLine(startMessage);
                searchedWord = Console.ReadLine();

                if (string.IsNullOrWhiteSpace(searchedWord) || !acceptedValues.Contains(searchedWord))
                {
                    Console.WriteLine(errorMessage);
                    searchedWord = string.Empty;
                }
            }
            while (string.IsNullOrWhiteSpace(searchedWord) || !acceptedValues.Contains(searchedWord));

            return searchedWord;
        }
    }
}