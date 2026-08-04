// vybe-test: csharp/csharp_scope_variables/out_variable_declared_inline_at_call_site
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

if(int.TryParse("42", out int n)) __P((n).ToString());
else __P((0).ToString());
__Check("42");
