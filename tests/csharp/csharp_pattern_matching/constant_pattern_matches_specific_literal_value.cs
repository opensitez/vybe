// vybe-test: csharp/csharp_pattern_matching/constant_pattern_matches_specific_literal_value
// origin: languages/csharp/tests/csharp/test_csharp_pattern_matching.rs

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

int x = 3;
string result = x switch { 1 => "one", 2 => "two", 3 => "three", _ => "other" };
__P((result).ToString());
__Check("three");
