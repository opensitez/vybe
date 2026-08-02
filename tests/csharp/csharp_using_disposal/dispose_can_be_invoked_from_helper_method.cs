// vybe-test: csharp/csharp_using_disposal/dispose_can_be_invoked_from_helper_method
// origin: languages/csharp/tests/csharp/test_csharp_using_disposal.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System; class Resource : IDisposable { public void Dispose() { __Check(("disposed").ToString(), "disposed"); } } void Close(IDisposable item) { item.Dispose(); } var resource = new Resource(); Close(resource);
