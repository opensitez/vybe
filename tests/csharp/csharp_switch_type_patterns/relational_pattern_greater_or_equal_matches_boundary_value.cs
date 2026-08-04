// vybe-test: csharp/csharp_switch_type_patterns/relational_pattern_greater_or_equal_matches_boundary_value
// origin: languages/csharp/tests/csharp/test_csharp_switch_type_patterns.rs

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

int score = 100;
string grade = score switch {
    >= 90 => "A",
    >= 80 => "B",
    _ => "C"
};
__P((grade).ToString());
__Check("A");
