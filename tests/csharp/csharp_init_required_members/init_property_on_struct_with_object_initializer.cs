// vybe-test: csharp/csharp_init_required_members/init_property_on_struct_with_object_initializer
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

struct Pair { public int A { get; init; } public int B { get; init; } }
var p = new Pair { A = 4, B = 6 };
__P((p.A + p.B).ToString());
__Check("10");
