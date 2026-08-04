// vybe-test: csharp/csharp_char_type_semantics/new_string_from_char_array_reconstructs_text
// origin: languages/csharp/tests/csharp/test_csharp_char_type_semantics.rs

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

char[] data = { 'h', 'i' };
__P((new string(data)).ToString());
__Check("hi");
