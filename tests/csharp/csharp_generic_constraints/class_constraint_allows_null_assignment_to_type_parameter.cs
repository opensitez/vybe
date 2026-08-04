// vybe-test: csharp/csharp_generic_constraints/class_constraint_allows_null_assignment_to_type_parameter
// origin: languages/csharp/tests/csharp/test_csharp_generic_constraints.rs

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

T AsNull<T>() where T : class => null;
__P((AsNull<string>() == null).ToString());
__Check("True");
