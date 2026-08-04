// vybe-test: csharp/csharp_init_required_members/init_property_two_instances_have_independent_values
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

class Slot { public int Id { get; init; } }
var a = new Slot { Id = 1 };
var b = new Slot { Id = 2 };
__P((a.Id).ToString()); __P((b.Id).ToString());
__Check("1\n2");
