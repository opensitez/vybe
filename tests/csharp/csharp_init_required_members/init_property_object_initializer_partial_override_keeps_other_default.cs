// vybe-test: csharp/csharp_init_required_members/init_property_object_initializer_partial_override_keeps_other_default
// origin: languages/csharp/tests/csharp/test_csharp_init_required_members.rs

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

class Pair { public int A { get; init; } = 1; public int B { get; init; } = 2; }
var p = new Pair { B = 9 };
__P((p.A).ToString()); __P((p.B).ToString());
__Check("1\n9");
