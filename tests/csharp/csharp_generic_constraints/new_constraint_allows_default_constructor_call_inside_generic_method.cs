// vybe-test: csharp/csharp_generic_constraints/new_constraint_allows_default_constructor_call_inside_generic_method
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

T Create<T>() where T : new() => new T();
class Widget { public int Value = 42; }
var w = Create<Widget>();
__P((w.Value).ToString());
__Check("42");
