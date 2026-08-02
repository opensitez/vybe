// vybe-test: csharp/csharp_generic_inference_calls/generic_method_constraint_struct_accepts_unboxed_value_type
// origin: languages/csharp/tests/csharp/test_csharp_generic_inference_calls.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int Scale<T>(T value) where T : struct {
    return 2 * (int)(object)value;
}
__Check((Scale(6)).ToString(), "12");
