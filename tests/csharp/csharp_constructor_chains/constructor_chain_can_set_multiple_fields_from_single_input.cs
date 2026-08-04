// vybe-test: csharp/csharp_constructor_chains/constructor_chain_can_set_multiple_fields_from_single_input
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

class Box { string left; string right; public Box(string value) : this(value, value.ToUpper()) { } public Box(string left, string right) { this.left = left; this.right = right; } public string Read() { return left + ":" + right; } } __P((new Box("a").Read()).ToString());
__Check("a:A");
