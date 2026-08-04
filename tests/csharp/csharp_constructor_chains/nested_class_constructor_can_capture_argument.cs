// vybe-test: csharp/csharp_constructor_chains/nested_class_constructor_can_capture_argument
// origin: languages/csharp/tests/csharp/test_csharp_constructor_chains.rs

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

class Outer { public class Inner { string name; public Inner(string name) { this.name = name; } public string Read() { return name; } } } __P((new Outer.Inner("inner").Read()).ToString());
__Check("inner");
