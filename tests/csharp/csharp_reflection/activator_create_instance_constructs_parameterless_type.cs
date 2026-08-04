// vybe-test: csharp/csharp_reflection/activator_create_instance_constructs_parameterless_type
// origin: languages/csharp/tests/csharp/test_csharp_reflection.rs

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

class Widget { public int Value = 42; }
var w = (Widget)System.Activator.CreateInstance(typeof(Widget));
__P((w.Value).ToString());
__Check("42");
