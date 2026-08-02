// vybe-test: csharp/csharp_using_disposal/using_statement_with_return_still_disposes_resource
// origin: languages/csharp/tests/csharp/test_csharp_using_disposal.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System; class Resource : IDisposable { public void Dispose() { __Check(("disposed").ToString(), "inside"); } } int Read() { using (var resource = new Resource()) { __Check(("inside").ToString(), "disposed"); return 5; } } __Check((Read()).ToString(), "5");
