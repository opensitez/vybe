// vybe-test: csharp/csharp_generic_inference_calls/generic_method_overload_resolution_prefers_specific_argument_types
// origin: languages/csharp/tests/csharp/test_csharp_generic_inference_calls.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

string Pick(int value) { return "int:" + value; }
string Pick(string value) { return "str:" + value; }
__Check((Pick(3)).ToString(), "int:3");
__Check((Pick("3")).ToString(), "str:3");
