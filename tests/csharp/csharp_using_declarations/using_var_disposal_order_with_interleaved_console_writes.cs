// vybe-test: csharp/csharp_using_declarations/using_var_disposal_order_with_interleaved_console_writes
// origin: languages/csharp/tests/csharp/test_csharp_using_declarations.rs

class R:System.IDisposable{string n;public R(string n){this.n=n;}public void Dispose(){Console.WriteLine("d:"+n);}}
using var a=new R("a"); Console.WriteLine("m1"); using var b=new R("b"); Console.WriteLine("m2");
