// vybe-test: csharp/csharp_using_declarations/using_var_with_break_in_loop_still_disposes
// origin: languages/csharp/tests/csharp/test_csharp_using_declarations.rs

class R:System.IDisposable{string n;public R(string n){this.n=n;}public void Dispose(){Console.WriteLine(n);}}
foreach(var n in new[]{1,2,3}){using var x=new R("b"); if(n==2) break; Console.WriteLine(n);} Console.WriteLine("end");
