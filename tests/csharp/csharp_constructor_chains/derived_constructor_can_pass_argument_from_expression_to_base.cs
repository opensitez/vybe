// vybe-test: csharp/csharp_constructor_chains/derived_constructor_can_pass_argument_from_expression_to_base
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

class Base { int value; public Base(int value) { this.value = value; } public int Read() { return value; } } class Child : Base { public Child(int value) : base(value + 1) { } } __P((new Child(4).Read()).ToString());
__Check("5");
