// vybe-test: csharp/csharp_constructor_chains/derived_constructor_can_pass_argument_from_expression_to_base
// origin: languages/csharp/tests/csharp/test_csharp_constructor_chains.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Base { int value; public Base(int value) { this.value = value; } public int Read() { return value; } } class Child : Base { public Child(int value) : base(value + 1) { } } __Check((new Child(4).Read()).ToString(), "5");
