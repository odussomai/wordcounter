using DocumentFormat.OpenXml.Packaging;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using System;
using System.IO;

namespace Word_Counter_For_Writers
{
    class Program
    {
        static void Main(string[] args)
        {
            var fileNames = Directory.GetFiles(@"C:\Users\User\Dropbox\stories\", "*.docx", SearchOption.AllDirectories);


            var totalWords = 0;
            foreach(var file in fileNames)
            {
                int words;
                try
                {
                    
                    using (var document = WordprocessingDocument.Open(file, false))
                    {
                        words = int.Parse(document.ExtendedFilePropertiesPart.Properties.Words?.Text);
                    }

                }
                catch
                {
                    continue;
                }


                totalWords += words;
            }

            Console.WriteLine(totalWords);
            
            Console.ReadLine();
        }

        public static IHostBuilder CreateHostBuilder(string[] args) => Host.CreateDefaultBuilder(args)
        .ConfigureServices((hostContext, services) =>
        {
            services.AddTransient<ICountWordsInADoc, CountWordsInADoc>();
        });
    }
}
