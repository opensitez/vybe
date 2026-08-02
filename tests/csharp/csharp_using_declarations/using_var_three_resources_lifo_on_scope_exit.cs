// vybe-test: csharp/csharp_using_declarations/using_var_three_resources_lifo_on_scope_exit
// origin: languages/csharp/tests/csharp/test_csharp_using_declarations.rs

class R:System.IDisposable{string n;public R(string n){this.n=n;}public void Dispose(){Console.WriteLine(n);}}
using var one=new R("1"); using var two=new R("2"); using var three=new R("3"); Console.WriteLine("mid");
