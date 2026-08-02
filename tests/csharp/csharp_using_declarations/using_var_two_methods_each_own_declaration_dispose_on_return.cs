// vybe-test: csharp/csharp_using_declarations/using_var_two_methods_each_own_declaration_dispose_on_return
// origin: languages/csharp/tests/csharp/test_csharp_using_declarations.rs

class R:System.IDisposable{string n;public R(string n){this.n=n;}public void Dispose(){Console.WriteLine(n);}}
void A(){using var x=new R("A"); Console.WriteLine("a");}
void B(){using var x=new R("B"); Console.WriteLine("b");}
A(); B();
