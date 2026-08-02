// vybe-test: csharp/csharp_nested_partial_types/nested_class_inside_generic_outer_type_uses_type_argument
// origin: languages/csharp/tests/csharp/test_csharp_nested_partial_types.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Box<T> {
    public class Wrapper {
        public T Value { get; set; }
    }
}
var wrapper = new Box<int>.Wrapper { Value = 9 };
__Check((wrapper.Value).ToString(), "9");
