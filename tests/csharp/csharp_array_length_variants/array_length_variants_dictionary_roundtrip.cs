// vybe-test: csharp/csharp_array_length_variants/array_length_variants_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_array_length_variants.rs

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

// array_length_variants
var map = new System.Collections.Generic.Dictionary<int, int>(); map[25] = 26; __P((map.ContainsKey(25) && map[25] == 26).ToString());
__Check("True");
