// vybe-test: csharp/csharp_using_declarations/using_var_lambda_invocation_disposes_after_lambda_returns
// origin: languages/csharp/tests/csharp/test_csharp_using_declarations.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class R:System.IDisposable{string n;public R(string n){this.n=n;}public void Dispose(){__Check((n).ToString(), "7");}}
System.Func<int> f=()=>{using var x=new R("lam"); return 7;}; __Check((f()).ToString(), "lam");
