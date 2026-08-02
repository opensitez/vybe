// vybe-test: csharp/csharp_constructor_chains/base_constructor_initializes_inherited_field
// origin: languages/csharp/tests/csharp/test_csharp_constructor_chains.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Base { protected string name; public Base(string name) { this.name = name; } public string Name() { return name; } } class Child : Base { public Child() : base("root") { } } __Check((new Child().Name()).ToString(), "root");
