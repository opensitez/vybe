// vybe-test: csharp/csharp_constructor_chains/constructor_can_accept_optional_parameter_value
// origin: languages/csharp/tests/csharp/test_csharp_constructor_chains.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Box { int value; public Box(int value = 6) { this.value = value; } public int Read() { return value; } } __Check((new Box().Read()).ToString(), "6");
