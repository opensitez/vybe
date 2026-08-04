// vybe-test: csharp/csharp_init_required_members/multiple_required_members_on_same_class
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

class Pair { public required int Left; public required int Right; }
var p = new Pair { Left = 3, Right = 9 };
__P((p.Left + p.Right).ToString());
__Check("12");
