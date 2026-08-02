// vybe-test: csharp/csharp_using_declarations/using_var_disposes_before_subsequent_using_var_peer_created_later
// origin: languages/csharp/tests/csharp/test_csharp_using_declarations.rs

class R:System.IDisposable{string n;public R(string n){this.n=n;}public void Dispose(){Console.WriteLine(n);}}
void Scope(){using var late=new R("late");} using var early=new R("early"); Scope(); Console.WriteLine("end");
