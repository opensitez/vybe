// vybe-test: csharp/csharp_generic_inference_calls/generic_method_constraint_class_allows_reference_type_members
// origin: languages/csharp/tests/csharp/test_csharp_generic_inference_calls.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string Describe<T>(T value) where T : class {
    return value == null ? "null" : value.ToString();
}
__Check((Describe("data")).ToString(), "data");
