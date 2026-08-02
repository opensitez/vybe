// vybe-test: csharp/csharp_using_declarations/using_var_empty_block_still_disposes_immediately
// origin: languages/csharp/tests/csharp/test_csharp_using_declarations.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class R:System.IDisposable{string n;public R(string n){this.n=n;}public void Dispose(){__Check((n).ToString(), "empty");}}
{using var x=new R("empty");} __Check(("next").ToString(), "next");
