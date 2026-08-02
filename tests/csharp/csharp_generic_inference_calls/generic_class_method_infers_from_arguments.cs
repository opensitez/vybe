// vybe-test: csharp/csharp_generic_inference_calls/generic_class_method_infers_from_arguments
// origin: languages/csharp/tests/csharp/test_csharp_generic_inference_calls.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Box<T> {
    public T Value;
    public Box(T value) { Value = value; }
    public T Get() { return Value; }
}
var numbers = new Box<int>(5);
__Check((numbers.Get()).ToString(), "5");
