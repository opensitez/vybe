// vybe-test: csharp/csharp_using_disposal/using_statement_disposes_multiple_resources_in_reverse_order
// origin: languages/csharp/tests/csharp/test_csharp_using_disposal.rs

using System; class Resource : IDisposable { string name; public Resource(string name) { this.name = name; } public void Dispose() { Console.WriteLine(name); } } using (var left = new Resource("left")) using (var right = new Resource("right")) { Console.WriteLine("body"); }
