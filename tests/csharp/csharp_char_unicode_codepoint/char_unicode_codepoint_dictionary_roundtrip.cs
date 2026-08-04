// vybe-test: csharp/csharp_char_unicode_codepoint/char_unicode_codepoint_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_char_unicode_codepoint.rs

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

// char_unicode_codepoint
var map = new System.Collections.Generic.Dictionary<int, int>(); map[22] = 23; __P((map.ContainsKey(22) && map[22] == 23).ToString());
__Check("True");
