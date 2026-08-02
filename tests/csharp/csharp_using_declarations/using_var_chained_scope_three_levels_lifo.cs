// vybe-test: csharp/csharp_using_declarations/using_var_chained_scope_three_levels_lifo
// origin: languages/csharp/tests/csharp/test_csharp_using_declarations.rs

class R:System.IDisposable{string n;public R(string n){this.n=n;}public void Dispose(){Console.WriteLine(n);}}
{using var l1=new R("l1"); {using var l2=new R("l2"); {using var l3=new R("l3"); Console.WriteLine("3");}} Console.WriteLine("2");} Console.WriteLine("1");
