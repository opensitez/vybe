// vybe-test: csharp/csharp_constructor_chains/this_constructor_chain_reuses_primary_overload
// origin: languages/csharp/tests/csharp/test_csharp_constructor_chains.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Box { int value; public Box() : this(9) { } public Box(int value) { this.value = value; } public int Read() { return value; } } __Check((new Box().Read()).ToString(), "9");
