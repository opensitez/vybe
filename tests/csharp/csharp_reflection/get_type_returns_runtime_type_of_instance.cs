// vybe-test: csharp/csharp_reflection/get_type_returns_runtime_type_of_instance
// origin: languages/csharp/tests/csharp/test_csharp_reflection.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

__Check((42.GetType().Name).ToString(), "Int32");
