// vybe-test: csharp/csharp_constructor_chains/object_initializer_runs_after_constructor_default_assignment
// origin: languages/csharp/tests/csharp/test_csharp_constructor_chains.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Box { public string Name { get; set; } public Box() { Name = "init"; } } var box = new Box { Name = "set" }; __Check((box.Name).ToString(), "set");
