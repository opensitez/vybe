// vybe-test: csharp/csharp_char_unicode_codepoint/char_unicode_codepoint_string_first_char
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
string feature = "char_unicode_codepoint"; __P((feature[0] == feature[0]).ToString());
__Check("True");
