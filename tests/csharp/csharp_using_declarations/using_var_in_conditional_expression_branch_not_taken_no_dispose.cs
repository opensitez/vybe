// vybe-test: csharp/csharp_using_declarations/using_var_in_conditional_expression_branch_not_taken_no_dispose
// origin: languages/csharp/tests/csharp/test_csharp_using_declarations.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class R:System.IDisposable{public static int N=0;public void Dispose(){N++;}}
bool ok=true; if(ok){using var x=new R(); __Check(("yes").ToString(), "yes");} else {using var y=new R();} __Check((R.N).ToString(), "1");
