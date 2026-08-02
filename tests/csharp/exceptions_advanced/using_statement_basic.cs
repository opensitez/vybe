// vybe-test: csharp/exceptions_advanced/using_statement_basic
// origin: languages/csharp/tests/csharp/test_exceptions_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Resource : IDisposable {
    public Resource() { __Check(("opened").ToString(), "opened"); }
    public void Dispose() { __Check(("disposed").ToString(), "using"); }
}
using (var r = new Resource()) {
    __Check(("using").ToString(), "disposed");
}
