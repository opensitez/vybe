// vybe-test: csharp/csharp_char_operations/char_is_whitespace_true_for_space_and_tab
// origin: languages/csharp/tests/csharp/test_csharp_char_operations.rs

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

__P((char.IsWhiteSpace(' ')).ToString()); __P((char.IsWhiteSpace('\t')).ToString());
__Check("True\nTrue");
