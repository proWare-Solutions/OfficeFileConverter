using CommandLine;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeFileConverter
{
  internal class Options
  {
    [Option('f', "file", Group ="File or directory", HelpText = "Path to a file with paths of files to convert.")]
    public string File { get; set; }

    [Option('d', "directory", Group = "File or directory", HelpText = "Path to a directory which is searched.")]
    public string Directory { get; set; }


    [Option('s', "settings", Required = true, HelpText = "Path to a settings file")]
    public string Settings { get; set; }

    [Option('l', "logdir", Required = false, HelpText = "Path to a log directory")]
    public string Logpath { get; set; }

    [Option('g', "debug", Required = false, HelpText = "Show debug information")]
    public bool Debug { get; set; }

    [Option('r', "remove", Required = false, HelpText = "Remove original", Default =false)]
    public bool RemoveOriginal { get; set; } = false;

    [Option('a', "allfiles", Required = false, HelpText = "Convert all files. If the parameter is not set, only macros will be converted", Default = false)]
    public bool AllFiles { get; set; } = false;

    [Option('c', "access", Required = false, HelpText = "Convert MS Access Files", Default = false)]
    public bool ConvertAccess { get; set; } = false;

    [Option('t', "tempdir", Required = false, HelpText = "Directory to store temporary files. Default %temp%")]
    public string TempDir { get; set; } = Environment.GetEnvironmentVariable("temp");

    public string CurrentRunTime { get; } = DateTime.Now.ToString("yyyy-MM-dd hh-mm-ss");

    public ConvertSettings ConvertSettings { get; set; }

    public int FilesProcessed { get; set; } = 0;

    public int FilesTotal { get; set; } = 0;

    public int DirectoriesTotal { get; set; } = 0;

    public string ActionStatusFile { get; set; }
  }
}
