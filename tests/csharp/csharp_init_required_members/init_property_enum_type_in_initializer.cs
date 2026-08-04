// vybe-test: csharp/csharp_init_required_members/init_property_enum_type_in_initializer
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

enum Level { Low, High }
class Job { public Level Tier { get; init; } = Level.Low; }
var j = new Job { Tier = Level.High };
__P((j.Tier).ToString());
__Check("High");
