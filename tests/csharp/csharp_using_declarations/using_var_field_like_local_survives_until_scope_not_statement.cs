// vybe-test: csharp/csharp_using_declarations/using_var_field_like_local_survives_until_scope_not_statement
// origin: languages/csharp/tests/csharp/test_csharp_using_declarations.rs

class R:System.IDisposable{string n;public R(string n){this.n=n;}public void Dispose(){Console.WriteLine(n);}}
using var hold=new R("hold"); for(int i=0;i<2;i++) Console.WriteLine(i);
