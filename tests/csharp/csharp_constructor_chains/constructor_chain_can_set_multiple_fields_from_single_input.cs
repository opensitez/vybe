// vybe-test: csharp/csharp_constructor_chains/constructor_chain_can_set_multiple_fields_from_single_input
// origin: languages/csharp/tests/csharp/test_csharp_constructor_chains.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Box { string left; string right; public Box(string value) : this(value, value.ToUpper()) { } public Box(string left, string right) { this.left = left; this.right = right; } public string Read() { return left + ":" + right; } } __Check((new Box("a").Read()).ToString(), "a:A");
