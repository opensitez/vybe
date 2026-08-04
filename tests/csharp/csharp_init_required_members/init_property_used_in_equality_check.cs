// vybe-test: csharp/csharp_init_required_members/init_property_used_in_equality_check
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

class Tag { public string Name { get; init; } = ""; }
var a = new Tag { Name = "x" };
var b = new Tag { Name = "x" };
__P((a.Name == b.Name).ToString());
__Check("True");
