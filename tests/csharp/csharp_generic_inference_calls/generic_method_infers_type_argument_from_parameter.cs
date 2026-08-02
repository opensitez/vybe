// vybe-test: csharp/csharp_generic_inference_calls/generic_method_infers_type_argument_from_parameter
// origin: languages/csharp/tests/csharp/test_csharp_generic_inference_calls.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

T Identity<T>(T value) { return value; }
__Check((Identity(42)).ToString(), "42");
__Check((Identity("text")).ToString(), "text");
