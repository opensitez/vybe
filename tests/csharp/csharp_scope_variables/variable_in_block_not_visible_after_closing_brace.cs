// vybe-test: csharp/csharp_scope_variables/variable_in_block_not_visible_after_closing_brace
// origin: languages/csharp/tests/csharp/test_csharp_scope_variables.rs

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

int outer = 1;
{
    int inner = 2;
    outer = inner;
}
__P((outer).ToString());
__Check("2");
