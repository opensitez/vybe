// vybe-test: csharp/csharp_using_declarations/using_var_in_switch_case_block_disposes_on_case_exit
// origin: languages/csharp/tests/csharp/test_csharp_using_declarations.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class R:System.IDisposable{string n;public R(string n){this.n=n;}public void Dispose(){__Check((n).ToString(), "case");}}
switch(1){case 1: using var x=new R("sw"); __Check(("case").ToString(), "sw"); break;} __Check(("after").ToString(), "after");
