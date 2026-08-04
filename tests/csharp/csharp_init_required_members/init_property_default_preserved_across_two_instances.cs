// vybe-test: csharp/csharp_init_required_members/init_property_default_preserved_across_two_instances
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

class Config { public int Retries { get; init; } = 3; }
var a = new Config();
var b = new Config { Retries = 1 };
__P((a.Retries).ToString()); __P((b.Retries).ToString());
__Check("3\n1");
