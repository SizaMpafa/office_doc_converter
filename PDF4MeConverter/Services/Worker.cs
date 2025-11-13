using Microsoft.Extensions.Hosting;
using System.IO;
using System.IO.Pipes;
using System.Threading;
using System.Threading.Tasks;

namespace PDF4MeConverter.Services;

public class Worker : BackgroundService
{
    private readonly OfficeConverter _converter;

    public Worker(OfficeConverter converter) { _converter = converter; }

    protected override async Task ExecuteAsync(CancellationToken stoppingToken)
    {
        Directory.CreateDirectory("input");
        Directory.CreateDirectory("output");

        using var watcher = new FileSystemWatcher("input", "*.docx") { EnableRaisingEvents = true };
        watcher.Created += async (s, e) =>
        {
            var output = Path.Combine("output", Path.GetFileNameWithoutExtension(e.Name) + ".pdf");
            await _converter.EnqueueAsync(e.FullPath, output);
        };
    _ = Task.Run(async () =>
        {
            while (!stoppingToken.IsCancellationRequested)
            {
                try
                {
                    using var pipeServer = new NamedPipeServerStream(
                        "PDF4MePipe",
                        PipeDirection.InOut,
                        1,
                        PipeTransmissionMode.Byte,
                        PipeOptions.Asynchronous);

                    await pipeServer.WaitForConnectionAsync(stoppingToken);

                    using var reader = new StreamReader(pipeServer);
                    using var writer = new StreamWriter(pipeServer) { AutoFlush = true };

                    while (!stoppingToken.IsCancellationRequested && pipeServer.IsConnected)
                    {
                        var request = await reader.ReadLineAsync();
                        if (request == null) break;

                        var parts = request.Split(':');
                        if (parts.Length == 3 && parts[0] == "convert")
                        {
                            await _converter.EnqueueAsync(parts[1], parts[2]);
                            await writer.WriteLineAsync("Conversion queued");
                        }
                        else
                        {
                            await writer.WriteLineAsync("ERROR: Invalid format");
                        }
                    }
                }
                catch (OperationCanceledException) { break; }
                catch (IOException)
                {
                    
                }
            }
        }, stoppingToken);
        await _converter.ProcessAsync(stoppingToken);
    }
}