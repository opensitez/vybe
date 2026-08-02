// vybe-test: csharp/csharp_generic_inference_calls/generic_method_new_constraint_allows_parameterless_construction
// origin: languages/csharp/tests/csharp/test_csharp_generic_inference_calls.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Widget { public int Size = 4; }
T Create<T>() where T : new() { return new T(); }
__Check((Create<Widget>().Size).ToString(), "4");
