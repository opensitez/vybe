// vybe-test: csharp/csharp_constructor_chains/constructor_overload_can_append_suffix_after_chain
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

class Box { string name; public Box(string name) { this.name = name; } public Box(string name, string suffix) : this(name) { this.name += suffix; } public string Read() { return name; } } __P((new Box("a", "b").Read()).ToString());
__Check("ab");
