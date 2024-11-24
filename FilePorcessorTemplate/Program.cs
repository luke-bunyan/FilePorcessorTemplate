using FilePorcessorTemplate.Helpers;

var fileTransformer = new FileTransformerService();

string path = Path.Combine(Environment.CurrentDirectory, "test_file.xlsx");

var output = fileTransformer.TransformFile(path);

Console.WriteLine("Hello, World!");