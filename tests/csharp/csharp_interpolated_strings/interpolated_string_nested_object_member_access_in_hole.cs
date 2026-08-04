// vybe-test: csharp/csharp_interpolated_strings/interpolated_string_nested_object_member_access_in_hole
// origin: languages/csharp/tests/csharp/test_csharp_interpolated_strings.rs

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

class Pair { public int A; public int B; }
var pair = new Pair { A = 2, B = 3 };
__P(($"{pair.A + pair.B}").ToString());
__Check("5");
