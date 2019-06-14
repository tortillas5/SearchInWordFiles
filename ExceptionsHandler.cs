// <copyright file="ExceptionsHandler.cs" company="Tortillas-Inc">
// Copy me, no rights reserved.
// </copyright>

namespace SearchInWordFiles
{
    using System;

    /// <summary>
    /// Class used to manage exceptions.
    /// </summary>
    public class ExceptionsHandler
    {
        /// <summary>
        /// Regroup every exception messages of an exception (Exception and inner exception).
        /// </summary>
        /// <param name="exception">Exception whose messages are going to be grouped.</param>
        /// <returns>String with grouped messages.</returns>
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