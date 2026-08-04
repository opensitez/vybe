// vybe-test: csharp/csharp_expression_increment_semantics/postfix_decrement_in_expression_uses_original_value
// origin: languages/csharp/tests/csharp/test_csharp_expression_increment_semantics.rs

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

int n = 3;
int total = n-- + n;
__P((total).ToString());
__P((n).ToString());
__Check("5\n2");
