// vybe-test: csharp/csharp_init_required_members/required_string_and_optional_init_property_together
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

class Profile { public required string User; public int Score { get; init; } = 0; }
var p = new Profile { User = "ada", Score = 100 };
__P((p.User).ToString()); __P((p.Score).ToString());
__Check("ada\n100");
