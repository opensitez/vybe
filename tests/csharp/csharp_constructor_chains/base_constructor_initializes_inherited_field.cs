// vybe-test: csharp/csharp_constructor_chains/base_constructor_initializes_inherited_field
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

class Base { protected string name; public Base(string name) { this.name = name; } public string Name() { return name; } } class Child : Base { public Child() : base("root") { } } __P((new Child().Name()).ToString());
__Check("root");
