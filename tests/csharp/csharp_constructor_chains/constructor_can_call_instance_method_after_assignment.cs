// vybe-test: csharp/csharp_constructor_chains/constructor_can_call_instance_method_after_assignment
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

class Box { int value; public Box(int value) { this.value = value; __P((Read()).ToString()); } public int Read() { return value; } } new Box(8);
__Check("8");
