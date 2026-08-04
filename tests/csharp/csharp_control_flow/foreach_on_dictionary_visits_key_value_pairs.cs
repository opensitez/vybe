// vybe-test: csharp/csharp_control_flow/foreach_on_dictionary_visits_key_value_pairs
// origin: languages/csharp/tests/csharp/test_csharp_control_flow.rs

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

var map = new System.Collections.Generic.Dictionary<string, int> { ["x"] = 1 };
int total = 0;
foreach (var pair in map) total += pair.Value;
__P((total).ToString());
__Check("1");
