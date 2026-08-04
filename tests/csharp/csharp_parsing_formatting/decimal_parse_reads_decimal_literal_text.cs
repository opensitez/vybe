// vybe-test: csharp/csharp_parsing_formatting/decimal_parse_reads_decimal_literal_text
// origin: languages/csharp/tests/csharp/test_csharp_parsing_formatting.rs

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

__P((decimal.Parse("7.25")).ToString());
__Check("7.25");
