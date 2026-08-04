// vybe-test: csharp/csharp_scope_variables/if_declaration_pattern_scopes_bound_variable_to_body
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

object o = "scoped";
if(o is string text)
    __P((text.Length).ToString());
__P(("done").ToString());
__Check("6\ndone");
