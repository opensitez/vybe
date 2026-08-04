// vybe-test: csharp/csharp_constructor_chains/this_constructor_chain_reuses_primary_overload
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

class Box { int value; public Box() : this(9) { } public Box(int value) { this.value = value; } public int Read() { return value; } } __P((new Box().Read()).ToString());
__Check("9");
