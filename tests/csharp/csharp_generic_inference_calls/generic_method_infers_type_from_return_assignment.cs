// vybe-test: csharp/csharp_generic_inference_calls/generic_method_infers_type_from_return_assignment
// origin: languages/csharp/tests/csharp/test_csharp_generic_inference_calls.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

T First<T>(T left, T right) { return left; }
string chosen = First("left", "right");
__Check((chosen).ToString(), "left");
