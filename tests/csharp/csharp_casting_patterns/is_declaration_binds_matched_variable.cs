// vybe-test: csharp/csharp_casting_patterns/is_declaration_binds_matched_variable
// origin: languages/csharp/tests/csharp/test_csharp_casting_patterns.rs

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

object o=42;
if(o is int n) __P((n*2).ToString());
__Check("84");
