// vybe-test: csharp/csharp_using_declarations/using_var_in_try_finally_finally_after_disposal
// origin: languages/csharp/tests/csharp/test_csharp_using_declarations.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class R:System.IDisposable{string n;public R(string n){this.n=n;}public void Dispose(){__Check((n).ToString(), "try");}}
try{using var x=new R("res"); __Check(("try").ToString(), "res");} finally{__Check(("fin").ToString(), "fin");}
