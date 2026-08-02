// vybe-test: csharp/csharp_constructor_chains/constructor_can_set_readonly_field_from_parameter
// origin: languages/csharp/tests/csharp/test_csharp_constructor_chains.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Box { readonly int value; public Box(int value) { this.value = value; } public int Read() { return value; } } __Check((new Box(7).Read()).ToString(), "7");
