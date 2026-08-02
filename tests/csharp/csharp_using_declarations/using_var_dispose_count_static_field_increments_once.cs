// vybe-test: csharp/csharp_using_declarations/using_var_dispose_count_static_field_increments_once
// origin: languages/csharp/tests/csharp/test_csharp_using_declarations.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class R:System.IDisposable{public static int N=0;public void Dispose(){N++; __Check((N).ToString(), "once");}}
using var x=new R(); __Check(("once").ToString(), "1");
