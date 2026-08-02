// vybe-test: csharp/csharp_static_type_behaviors/instance_method_calls_static_helper_on_same_type
// origin: languages/csharp/tests/csharp/test_csharp_static_type_behaviors.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Converter {
    public static int Double(int value) { return value * 2; }
    public int Convert(int value) { return Double(value) + 1; }
}
var converter = new Converter();
__Check((converter.Convert(5)).ToString(), "11");
