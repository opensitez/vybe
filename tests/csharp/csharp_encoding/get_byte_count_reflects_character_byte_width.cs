// vybe-test: csharp/csharp_encoding/get_byte_count_reflects_character_byte_width
// origin: languages/csharp/tests/csharp/test_csharp_encoding.rs

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

int n = System.Text.Encoding.UTF8.GetByteCount("café");
__P((n > 4).ToString());
__Check("True");
