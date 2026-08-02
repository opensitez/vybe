// vybe-test: csharp/csharp_using_disposal/using_declaration_disposes_at_end_of_scope
// origin: languages/csharp/tests/csharp/test_csharp_using_disposal.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System; class Resource : IDisposable { public void Dispose() { __Check(("disposed").ToString(), "body"); } } using var resource = new Resource(); __Check(("body").ToString(), "disposed");
