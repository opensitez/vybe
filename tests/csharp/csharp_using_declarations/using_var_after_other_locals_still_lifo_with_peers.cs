// vybe-test: csharp/csharp_using_declarations/using_var_after_other_locals_still_lifo_with_peers
// origin: languages/csharp/tests/csharp/test_csharp_using_declarations.rs

class R:System.IDisposable{string n;public R(string n){this.n=n;}public void Dispose(){Console.WriteLine(n);}}
int count=0; using var a=new R("a"); count++; using var b=new R("b"); count++; Console.WriteLine(count);
