// <copyright file="GestionReadWriteConsole.cs" company="GFI-GEOSPHERE">
// Copyright (c) 2018 All Rights Reserved
// </copyright>

namespace SearchInWordFiles
{
    using System;
    using System.Collections.Generic;
    using System.Linq;

    /// <summary>
    /// Classe qui propose des méthodes de gestion de lecture / écriture dans la console, pour faciliter l'interaction avec l'utilisateur.
    /// </summary>
    public static class GestionReadWriteConsole
    {
        /// <summary>
        /// Demande à l'utilisateur de saisir du texte. Vérifie que le texte rentré est correct.
        /// </summary>
        /// <param name="startMessage">Message de départ, indiquant à l'utilisateur ce qu'il faut rentrer comme paramètre.</param>
        /// <param name="errorMessage">Message d'erreur si l’utilisateur rentre une valeur inexacte dans la console.</param>
        /// <returns>Retourne le texte saisi par l'utilisateur après l'avoir vérifié.</returns>
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
        /// Demande à l'utilisateur de saisir du texte. Vérifie que le texte rentré est correct. Vérifie que le texte rentré correspond à une liste de valeurs prédéfinies.
        /// </summary>
        /// <param name="startMessage">Message de départ, indiquant à l'utilisateur ce qu'il faut rentrer comme paramètre.</param>
        /// <param name="errorMessage">Message d'erreur si l’utilisateur rentre une valeur inexacte dans la console.</param>
        /// <param name="acceptedValues">Liste des valeurs acceptées dans la saisie de l'utilisateur.</param>
        /// <returns>Retourne le texte saisi par l'utilisateur après l'avoir vérifié.</returns>
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