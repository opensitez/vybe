// vybe-test: csharp/csharp_readonly_members/readonly_field_set_in_constructor_cannot_change_after
// origin: languages/csharp/tests/csharp/test_csharp_readonly_members.rs

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

class Immutable{public readonly int Value; public Immutable(int v){Value=v;}}
var obj=new Immutable(42);
__P((obj.Value).ToString());
__Check("42");
