// vybe-test: csharp/csharp_using_declarations/using_var_in_for_loop_body_disposes_each_iteration
// origin: languages/csharp/tests/csharp/test_csharp_using_declarations.rs

class R:System.IDisposable{string n;public R(string n){this.n=n;}public void Dispose(){Console.WriteLine(n);}}
for(int i=0;i<2;i++){using var x=new R("f"+i); Console.WriteLine("iter");} Console.WriteLine("done");
