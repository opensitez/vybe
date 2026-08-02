// vybe-test: csharp/csharp_using_declarations/using_var_disposes_after_following_statement
// origin: languages/csharp/tests/csharp/test_csharp_using_declarations.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class R:System.IDisposable{string n;public R(string n){this.n=n;}public void Dispose(){__Check((n).ToString(), "body");}}
using var x=new R("x"); __Check(("body").ToString(), "x");
