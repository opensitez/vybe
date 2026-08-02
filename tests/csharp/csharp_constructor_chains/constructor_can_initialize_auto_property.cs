// vybe-test: csharp/csharp_constructor_chains/constructor_can_initialize_auto_property
// origin: languages/csharp/tests/csharp/test_csharp_constructor_chains.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Box { public string Name { get; } public Box(string name) { Name = name; } } __Check((new Box("pkg").Name).ToString(), "pkg");
