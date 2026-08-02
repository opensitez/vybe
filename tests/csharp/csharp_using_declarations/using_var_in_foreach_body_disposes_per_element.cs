// vybe-test: csharp/csharp_using_declarations/using_var_in_foreach_body_disposes_per_element
// origin: languages/csharp/tests/csharp/test_csharp_using_declarations.rs

class R:System.IDisposable{string n;public R(string n){this.n=n;}public void Dispose(){Console.WriteLine(n);}}
foreach(var n in new[]{1,2}){using var x=new R("e"+n); Console.WriteLine("each");} Console.WriteLine("all");
