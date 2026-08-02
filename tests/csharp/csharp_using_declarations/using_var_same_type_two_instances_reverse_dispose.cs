// vybe-test: csharp/csharp_using_declarations/using_var_same_type_two_instances_reverse_dispose
// origin: languages/csharp/tests/csharp/test_csharp_using_declarations.rs

class R:System.IDisposable{int id;public R(int id){this.id=id;}public void Dispose(){Console.WriteLine(id);}}
using var first=new R(1); using var second=new R(2); Console.WriteLine("pair");
