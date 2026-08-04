// vybe-test: csharp/csharp_string_ops_advanced/string_to_char_array_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_string_ops_advanced.rs

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

char[] chars="abc".ToCharArray();
__P((new string(chars)).ToString());
__Check("abc");
