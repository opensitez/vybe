// vybe-test: csharp/csharp_using_declarations/using_var_guard_clause_return_path_disposes
// origin: languages/csharp/tests/csharp/test_csharp_using_declarations.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class R:System.IDisposable{string n;public R(string n){this.n=n;}public void Dispose(){__Check((n).ToString(), "pick");}}
string Pick(int v){using var x=new R("pick"); if(v<0) return "neg"; return "pos";} __Check((Pick(-1)).ToString(), "neg");
