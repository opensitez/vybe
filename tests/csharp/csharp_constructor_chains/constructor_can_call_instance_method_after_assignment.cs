// vybe-test: csharp/csharp_constructor_chains/constructor_can_call_instance_method_after_assignment
// origin: languages/csharp/tests/csharp/test_csharp_constructor_chains.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Box { int value; public Box(int value) { this.value = value; __Check((Read()).ToString(), "8"); } public int Read() { return value; } } new Box(8);
