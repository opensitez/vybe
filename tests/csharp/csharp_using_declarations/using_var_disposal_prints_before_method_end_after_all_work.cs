// vybe-test: csharp/csharp_using_declarations/using_var_disposal_prints_before_method_end_after_all_work
// origin: languages/csharp/tests/csharp/test_csharp_using_declarations.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class R:System.IDisposable{string n;public R(string n){this.n=n;}public void Dispose(){__Check((n).ToString(), "work");}}
void Run(){using var x=new R("run"); __Check(("work").ToString(), "run");} Run(); __Check(("after").ToString(), "after");
