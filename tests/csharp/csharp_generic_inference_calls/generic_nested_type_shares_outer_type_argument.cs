// vybe-test: csharp/csharp_generic_inference_calls/generic_nested_type_shares_outer_type_argument
// origin: languages/csharp/tests/csharp/test_csharp_generic_inference_calls.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Outer<T> {
    public class Inner {
        public T Value;
    }
    public Inner Build(T value) {
        return new Inner { Value = value };
    }
}
var built = new Outer<string>().Build("nested");
__Check((built.Value).ToString(), "nested");
