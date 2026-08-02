// vybe-test: csharp/csharp_nested_partial_types/nested_class_reads_outer_constant_value
// origin: languages/csharp/tests/csharp/test_csharp_nested_partial_types.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Outer {
    public const string Prefix = "outer";
    public class Inner {
        public string Read() { return Prefix + "/inner"; }
    }
}
__Check((new Outer.Inner().Read()).ToString(), "outer/inner");
