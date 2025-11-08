using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.DependencyInjection;
using PDF4MeConverter.Services;

namespace PDF4MeConverter;
public class Program
{
    public static async Task Main(string[] args)
    {
        var builder = Host.CreateApplicationBuilder(args);
        builder.Services.AddSingleton<OfficeConverter>();
        builder.Services.AddHostedService<Worker>();
        var host = builder.Build();
        await host.RunAsync();
    }
}    