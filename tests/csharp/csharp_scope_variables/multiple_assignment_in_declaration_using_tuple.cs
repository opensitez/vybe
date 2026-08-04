// vybe-test: csharp/csharp_scope_variables/multiple_assignment_in_declaration_using_tuple
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

var (a, b) = (3, 7);
__P((a).ToString()); __P((b).ToString());
__Check("3\n7");
