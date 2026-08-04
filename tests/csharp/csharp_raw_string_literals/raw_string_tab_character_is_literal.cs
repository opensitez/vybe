// vybe-test: csharp/csharp_raw_string_literals/raw_string_tab_character_is_literal
// origin: languages/csharp/tests/csharp/test_csharp_raw_string_literals.rs

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

string text="""a	b"""; __P((text.Length).ToString());
__Check("3");
