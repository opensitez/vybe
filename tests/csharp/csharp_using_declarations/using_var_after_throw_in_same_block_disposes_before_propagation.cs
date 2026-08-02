// vybe-test: csharp/csharp_using_declarations/using_var_after_throw_in_same_block_disposes_before_propagation
// origin: languages/csharp/tests/csharp/test_csharp_using_declarations.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class R:System.IDisposable{string n;public R(string n){this.n=n;}public void Dispose(){__Check((n).ToString(), "x");}}
try{using var x=new R("x"); throw new System.InvalidOperationException();} catch{__Check(("handled").ToString(), "handled");}
