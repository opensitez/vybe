// vybe-test: csharp/csharp_string_builder_usage/string_builder_usage_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_string_builder_usage.rs

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

// string_builder_usage
var map = new System.Collections.Generic.Dictionary<int, int>(); map[20] = 21; __P((map.ContainsKey(20) && map[20] == 21).ToString());
__Check("True");
