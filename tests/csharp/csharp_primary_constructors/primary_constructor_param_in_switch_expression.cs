// vybe-test: csharp/csharp_primary_constructors/primary_constructor_param_in_switch_expression
// origin: languages/csharp/tests/csharp/test_csharp_primary_constructors.rs

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

class Mode(int code) { public string Name() => code switch { 1 => "a", 2 => "b", _ => "x" }; }
__P((new Mode(2).Name()).ToString());
__Check("b");
