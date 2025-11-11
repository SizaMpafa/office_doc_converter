using Microsoft.Extensions.Hosting;
using System.IO;
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

        await _converter.ProcessAsync(stoppingToken);
    }
}