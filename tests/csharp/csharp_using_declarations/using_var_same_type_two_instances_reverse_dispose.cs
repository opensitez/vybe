// vybe-test: csharp/csharp_using_declarations/using_var_same_type_two_instances_reverse_dispose
// origin: languages/csharp/tests/csharp/test_csharp_using_declarations.rs

string __buf = "";

void __P(string s) {
    __buf = __buf + s + "\n";
}

void __Pr(string s) {
    __buf = __buf + s;
}

// The final WriteLine contributes a trailing newline that the expected line
// vector never carried, so BOTH forms are accepted.
void __Check(string want) {
    if (__buf != want && __buf != want + "\n") {
        Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
        throw new Exception("assertion failed");
    }
}

class R:System.IDisposable{int id;public R(int id){this.id=id;}public void Dispose(){__P((id).ToString());}}
using var first=new R(1); using var second=new R(2); __P(("pair").ToString());
__Check("pair\n2\n1");
