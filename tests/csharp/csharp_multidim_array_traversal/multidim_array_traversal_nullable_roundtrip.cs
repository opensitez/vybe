// vybe-test: csharp/csharp_multidim_array_traversal/multidim_array_traversal_nullable_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_multidim_array_traversal.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// multidim_array_traversal
int? maybe = 29; __Check((maybe.HasValue && maybe.Value == 29).ToString(), "True");
