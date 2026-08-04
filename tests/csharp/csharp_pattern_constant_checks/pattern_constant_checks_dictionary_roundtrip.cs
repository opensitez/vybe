// vybe-test: csharp/csharp_pattern_constant_checks/pattern_constant_checks_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_pattern_constant_checks.rs

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

// pattern_constant_checks
var map = new System.Collections.Generic.Dictionary<int, int>(); map[40] = 41; __P((map.ContainsKey(40) && map[40] == 41).ToString());
__Check("True");
