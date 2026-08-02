// vybe-test: csharp/csharp_reflection_activation/get_nested_type_finds_declared_inner_class
// origin: languages/csharp/tests/csharp/test_csharp_reflection_activation.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Outer { public class Inner { } } __Check((typeof(Outer).GetNestedType("Inner") != null).ToString(), "True");
