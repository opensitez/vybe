// vybe-test: csharp/csharp_using_declarations/using_var_with_early_return_still_disposes
// origin: languages/csharp/tests/csharp/test_csharp_using_declarations.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class R:System.IDisposable{string n;public R(string n){this.n=n;}public void Dispose(){__Check((n).ToString(), "go");}}
int Go(bool stop){using var x=new R("go"); if(stop) return 1; return 2;} __Check((Go(true)).ToString(), "1");
