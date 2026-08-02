// vybe-test: csharp/csharp_using_declarations/using_var_nested_block_inner_then_outer_disposal
// origin: languages/csharp/tests/csharp/test_csharp_using_declarations.rs

class R:System.IDisposable{string n;public R(string n){this.n=n;}public void Dispose(){Console.WriteLine(n);}}
using var outer=new R("outer");
{using var inner=new R("inner"); Console.WriteLine("nest");}
Console.WriteLine("flat");
