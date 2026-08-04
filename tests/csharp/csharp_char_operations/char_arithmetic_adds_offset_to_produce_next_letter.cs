// vybe-test: csharp/csharp_char_operations/char_arithmetic_adds_offset_to_produce_next_letter
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

char c = (char)('A' + 2); __P((c).ToString());
__Check("C");
