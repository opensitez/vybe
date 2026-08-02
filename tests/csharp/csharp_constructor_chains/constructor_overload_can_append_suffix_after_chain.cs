// vybe-test: csharp/csharp_constructor_chains/constructor_overload_can_append_suffix_after_chain
// origin: languages/csharp/tests/csharp/test_csharp_constructor_chains.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Box { string name; public Box(string name) { this.name = name; } public Box(string name, string suffix) : this(name) { this.name += suffix; } public string Read() { return name; } } __Check((new Box("a", "b").Read()).ToString(), "ab");
