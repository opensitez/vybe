// vybe-test: csharp/csharp_constructor_chains/base_constructor_and_override_dispatch_can_coexist_after_construction
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

class Base { protected string prefix; public Base(string prefix) { this.prefix = prefix; } public virtual string Read() { return prefix; } } class Child : Base { public Child() : base("x") { } public override string Read() { return prefix + "y"; } } __P((new Child().Read()).ToString());
__Check("xy");
