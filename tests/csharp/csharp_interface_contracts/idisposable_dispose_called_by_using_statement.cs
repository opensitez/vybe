// vybe-test: csharp/csharp_interface_contracts/idisposable_dispose_called_by_using_statement
// origin: languages/csharp/tests/csharp/test_csharp_interface_contracts.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Resource : System.IDisposable {
    public static int Disposed = 0;
    public void Dispose() => Disposed++;
}
using(var r = new Resource()) { }
__Check((Resource.Disposed).ToString(), "1");
