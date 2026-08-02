// vybe-test: csharp/csharp_using_declarations/using_var_in_while_loop_body_disposes_each_iteration
// origin: languages/csharp/tests/csharp/test_csharp_using_declarations.rs

class R:System.IDisposable{string n;public R(string n){this.n=n;}public void Dispose(){Console.WriteLine(n);}}
int i=0; while(i<2){using var x=new R(i.ToString()); Console.WriteLine("loop"); i++;} Console.WriteLine("exit");
