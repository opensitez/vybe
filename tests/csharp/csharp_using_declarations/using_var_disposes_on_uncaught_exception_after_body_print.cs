// vybe-test: csharp/csharp_using_declarations/using_var_disposes_on_uncaught_exception_after_body_print
// origin: languages/csharp/tests/csharp/test_csharp_using_declarations.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class R:System.IDisposable{string n;public R(string n){this.n=n;}public void Dispose(){__Check((n).ToString(), "body");}}
try{using var x=new R("boom"); __Check(("body").ToString(), "boom"); throw new System.Exception();} catch{__Check(("caught").ToString(), "caught");}
