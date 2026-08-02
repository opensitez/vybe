// vybe-test: csharp/csharp_using_declarations/using_var_in_if_block_disposes_before_if_ends
// origin: languages/csharp/tests/csharp/test_csharp_using_declarations.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class R:System.IDisposable{string n;public R(string n){this.n=n;}public void Dispose(){__Check((n).ToString(), "then");}}
if(true){using var x=new R("if"); __Check(("then").ToString(), "if");} __Check(("after").ToString(), "after");
