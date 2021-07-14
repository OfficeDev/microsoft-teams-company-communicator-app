// <copyright file="Program.cs" company="Microsoft">
// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.
// </copyright>

namespace Microsoft.Teams.Apps.CompanyCommunicator
{
    using Microsoft.AspNetCore.Hosting;
    using Microsoft.Extensions.DependencyInjection;
    using Microsoft.Extensions.Hosting;

    /// <summary>
    /// Program class of the company communicator application.
    /// </summary>
    public class Program
    {
        /// <summary>
        /// Main function of the company communicator application.
        /// It builds a web host, then launches the company communicator into it.
        /// </summary>
        /// <param name="args">Arguments passed in to the function.</param>
        public static void Main(string[] args)
        {
            CreateHostBuilder(args).Build().Run();
        }

        /// <summary>
        /// Create the web host builder.
        /// </summary>
        /// <param name="args">Arguments passed into the main function.</param>
        /// <returns>A web host builder instance.</returns>
        public static IHostBuilder CreateHostBuilder(string[] args) =>
           Host.CreateDefaultBuilder(args)
               .ConfigureWebHostDefaults(webBuilder =>
               {
                   webBuilder.UseStartup<Startup>();
               })
               .ConfigureServices(services =>
               {
                services.AddHostedService<SendMessageScheduler>();
               });
    }
}