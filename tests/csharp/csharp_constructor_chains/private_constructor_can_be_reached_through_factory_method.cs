// vybe-test: csharp/csharp_constructor_chains/private_constructor_can_be_reached_through_factory_method
// origin: languages/csharp/tests/csharp/test_csharp_constructor_chains.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Box { string name; private Box(string name) { this.name = name; } public static Box Create() { return new Box("made"); } public string Read() { return name; } } __Check((Box.Create().Read()).ToString(), "made");
