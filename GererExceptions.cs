// <copyright file="GererExceptions.cs" company="GFI-GEOSPHERE">
// Copyright (c) 2018 All Rights Reserved
// </copyright>

namespace SearchInWordFiles
{
    using System;

    public class GererExceptions
    {
        /// <summary>
        /// Concatène les message d'une exception et de ses potentiel sous exceptions.
        /// </summary>
        /// <param name="exception">Exception dont on veut concaténer les messages.</param>
        /// <returns>Message de toutes les exceptions, concaténés.</returns>
        public static string ConcatException(Exception exception)
        {
            string message = string.Empty;
            Exception ex = exception;

            do
            {
                message += ex.Message + "\n";
                ex = ex.InnerException;
            }
            while (ex != null);

            return message;
        }
    }
}