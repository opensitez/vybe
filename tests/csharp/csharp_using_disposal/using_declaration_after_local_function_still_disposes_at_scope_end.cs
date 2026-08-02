// vybe-test: csharp/csharp_using_disposal/using_declaration_after_local_function_still_disposes_at_scope_end
// origin: languages/csharp/tests/csharp/test_csharp_using_disposal.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System; class Resource : IDisposable { public void Dispose() { __Check(("disposed").ToString(), "disposed"); } } string Read() { using var resource = new Resource(); return "ok"; } __Check((Read()).ToString(), "ok");
