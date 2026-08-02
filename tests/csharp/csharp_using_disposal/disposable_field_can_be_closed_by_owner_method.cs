// vybe-test: csharp/csharp_using_disposal/disposable_field_can_be_closed_by_owner_method
// origin: languages/csharp/tests/csharp/test_csharp_using_disposal.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System; class Resource : IDisposable { public void Dispose() { __Check(("disposed").ToString(), "disposed"); } } class Owner { Resource resource = new Resource(); public void Close() { resource.Dispose(); } } new Owner().Close();
