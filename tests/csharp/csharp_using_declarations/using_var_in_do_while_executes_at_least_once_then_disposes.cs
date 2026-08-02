// vybe-test: csharp/csharp_using_declarations/using_var_in_do_while_executes_at_least_once_then_disposes
// origin: languages/csharp/tests/csharp/test_csharp_using_declarations.rs

class R:System.IDisposable{string n;public R(string n){this.n=n;}public void Dispose(){Console.WriteLine(n);}}
int n=0; do{using var x=new R("do"); Console.WriteLine("once"); n++;} while(n<1); Console.WriteLine("fin");
