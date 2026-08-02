// vybe-test: csharp/csharp_using_declarations/using_var_in_else_block_disposes_on_else_exit
// origin: languages/csharp/tests/csharp/test_csharp_using_declarations.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class R:System.IDisposable{string n;public R(string n){this.n=n;}public void Dispose(){__Check((n).ToString(), "branch");}}
if(false){} else {using var x=new R("else"); __Check(("branch").ToString(), "else");} __Check(("end").ToString(), "end");
