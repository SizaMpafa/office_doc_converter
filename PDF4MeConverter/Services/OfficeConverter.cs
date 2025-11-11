using System;
using System.IO;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Word;
using System.Threading.Channels;
using System.Threading;
using System.Threading.Tasks;
using Polly;

namespace PDF4MeConverter.Services;

public class OfficeConverter
{
    private readonly Channel<(string input, string output)> _queue = Channel.CreateUnbounded<(string, string)>();
    private readonly SemaphoreSlim _semaphore = new(1, 1);

    public async System.Threading.Tasks.Task EnqueueAsync(string input, string output) => await _queue.Writer.WriteAsync((input, output));

    public async System.Threading.Tasks.Task ProcessAsync(CancellationToken token)
    {
        while (!token.IsCancellationRequested)
        {
            if (await _queue.Reader.WaitToReadAsync(token))
            {
                var job = await _queue.Reader.ReadAsync(token);
                await _semaphore.WaitAsync(token);
                try
                {
                    await Policy.Handle<Exception>().WaitAndRetryAsync(3, i => TimeSpan.FromSeconds(i * 2))
                        .ExecuteAsync(() => ConvertToPdfAsync(job.input, job.output));
                }
                catch (Exception ex) { Console.WriteLine(ex.Message); }
                finally { _semaphore.Release(); }
            }
        }
        
    }

    private static async System.Threading.Tasks.Task ConvertToPdfAsync(string input, string output)
    {
        Application app = null!;
        Document doc = null!;
        try
        {
            app = new Application { Visible = false };
            doc = app.Documents.Open(input);
            doc.ExportAsFixedFormat(output, WdExportFormat.wdExportFormatPDF);
        }
        finally
        {
            doc?.Close(WdSaveOptions.wdDoNotSaveChanges);

            if (doc != null)
            {
                                doc.Close(WdSaveOptions.wdDoNotSaveChanges);
                #pragma warning disable CA1416
                                Marshal.ReleaseComObject(doc);
                #pragma warning restore CA1416
                            }
                            if (app != null)
                            {
                                app.Quit();
                #pragma warning disable CA1416
                                Marshal.ReleaseComObject(app);
                #pragma warning restore CA1416
            }
            GC.Collect();
            GC.WaitForPendingFinalizers();
            await System.Threading.Tasks.Task.CompletedTask;
        }
    }
}