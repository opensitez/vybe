// vybe-test: csharp/csharp_using_declarations/using_var_two_in_same_scope_dispose_reverse_order
// origin: languages/csharp/tests/csharp/test_csharp_using_declarations.rs

class R:System.IDisposable{string n;public R(string n){this.n=n;}public void Dispose(){Console.WriteLine(n);}}
using var a=new R("a"); using var b=new R("b"); Console.WriteLine("done");
