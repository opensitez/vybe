// vybe-test: csharp/csharp_switch_type_patterns/switch_on_string_matches_exact_literal_case
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

string Pick(string key) {
    switch (key) {
        case "go": return "G";
        case "stop": return "S";
        default: return "?";
    }
}
__P((Pick("go")).ToString());
__Check("G");
