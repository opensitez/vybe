// vybe-test: csharp/csharp_using_declarations/using_var_multiple_returns_same_scope_one_dispose
// origin: languages/csharp/tests/csharp/test_csharp_using_declarations.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class R:System.IDisposable{string n;public R(string n){this.n=n;}public void Dispose(){__Check((n).ToString(), "f");}}
int F(int n){using var x=new R("f"); if(n==0) return 0; if(n==1) return 1; return 2;} __Check((F(1)).ToString(), "1");
