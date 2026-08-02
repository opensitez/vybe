// vybe-test: csharp/csharp_using_disposal/using_statement_runs_dispose_when_exception_is_caught
// origin: languages/csharp/tests/csharp/test_csharp_using_disposal.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System; class Resource : IDisposable { public void Dispose() { __Check(("disposed").ToString(), "body"); } } try { using (var resource = new Resource()) { __Check(("body").ToString(), "disposed"); throw new Exception("boom"); } } catch (Exception) { __Check(("caught").ToString(), "caught"); }
