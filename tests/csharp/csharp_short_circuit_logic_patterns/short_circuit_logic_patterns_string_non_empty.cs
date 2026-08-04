// vybe-test: csharp/csharp_short_circuit_logic_patterns/short_circuit_logic_patterns_string_non_empty
// origin: languages/csharp/tests/csharp/test_csharp_short_circuit_logic_patterns.rs

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

// short_circuit_logic_patterns
string feature = "short_circuit_logic_patterns"; __P((feature.Length > 0).ToString());
__Check("True");
