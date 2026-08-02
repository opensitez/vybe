// vybe-test: csharp/csharp_using_declarations/using_var_in_local_function_nested_disposes_on_fn_exit
// origin: languages/csharp/tests/csharp/test_csharp_using_declarations.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class R:System.IDisposable{string n;public R(string n){this.n=n;}public void Dispose(){__Check((n).ToString(), "fn");}}
void Outer(){void Inner(){using var x=new R("in"); __Check(("fn").ToString(), "in");} Inner(); __Check(("out").ToString(), "out");} Outer();
