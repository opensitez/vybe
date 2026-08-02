// vybe-test: csharp/csharp_using_declarations/using_var_disposal_runs_after_all_prior_writes
// origin: languages/csharp/tests/csharp/test_csharp_using_declarations.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class R:System.IDisposable{string n;public R(string n){this.n=n;}public void Dispose(){__Check((n).ToString(), "1");}}
using var a=new R("last"); __Check(("1").ToString(), "2"); __Check(("2").ToString(), "last");
