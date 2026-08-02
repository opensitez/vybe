// vybe-test: csharp/csharp_constructor_chains/base_constructor_and_override_dispatch_can_coexist_after_construction
// origin: languages/csharp/tests/csharp/test_csharp_constructor_chains.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Base { protected string prefix; public Base(string prefix) { this.prefix = prefix; } public virtual string Read() { return prefix; } } class Child : Base { public Child() : base("x") { } public override string Read() { return prefix + "y"; } } __Check((new Child().Read()).ToString(), "xy");
