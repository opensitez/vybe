// vybe-test: csharp/csharp_constructor_chains/nested_class_constructor_can_capture_argument
// origin: languages/csharp/tests/csharp/test_csharp_constructor_chains.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Outer { public class Inner { string name; public Inner(string name) { this.name = name; } public string Read() { return name; } } } __Check((new Outer.Inner("inner").Read()).ToString(), "inner");
