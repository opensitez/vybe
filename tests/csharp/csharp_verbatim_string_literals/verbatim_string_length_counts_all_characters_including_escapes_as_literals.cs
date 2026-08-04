// vybe-test: csharp/csharp_verbatim_string_literals/verbatim_string_length_counts_all_characters_including_escapes_as_literals
// origin: languages/csharp/tests/csharp/test_csharp_verbatim_string_literals.rs

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

__P((@"\".Length).ToString());
__Check("2");
