// vybe-test: csharp/csharp_using_disposal/nested_using_declarations_dispose_in_reverse_order
// origin: languages/csharp/tests/csharp/test_csharp_using_disposal.rs

using System; class Resource : IDisposable { string name; public Resource(string name) { this.name = name; } public void Dispose() { Console.WriteLine(name); } } using var first = new Resource("first"); using var second = new Resource("second"); Console.WriteLine("done");
