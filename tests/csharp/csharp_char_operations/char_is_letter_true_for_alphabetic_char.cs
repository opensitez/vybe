// vybe-test: csharp/csharp_char_operations/char_is_letter_true_for_alphabetic_char
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

__P((char.IsLetter('a')).ToString());
__Check("True");
