// vybe-test: csharp/csharp_using_declarations/using_var_before_return_from_local_function
// origin: languages/csharp/tests/csharp/test_csharp_using_declarations.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class R:System.IDisposable{string n;public R(string n){this.n=n;}public void Dispose(){__Check((n).ToString(), "fn");}}
string Read(){using var x=new R("fn"); return "ok";} __Check((Read()).ToString(), "ok");
