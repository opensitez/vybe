// vybe-test: csharp/csharp_char_type_semantics/char_literal_single_quote_denotes_code_unit
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

char letter = 'A'; __P((letter).ToString());
__Check("A");
