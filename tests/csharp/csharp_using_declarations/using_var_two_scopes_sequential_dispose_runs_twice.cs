// vybe-test: csharp/csharp_using_declarations/using_var_two_scopes_sequential_dispose_runs_twice
// origin: languages/csharp/tests/csharp/test_csharp_using_declarations.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class R:System.IDisposable{public static int N=0;public void Dispose(){N++;}}
{using var x=new R();} {using var y=new R();} __Check((R.N).ToString(), "2");
