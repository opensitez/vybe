// vybe-test: csharp/csharp_array_copy_behavior/array_copy_behavior_nullable_fallback
// origin: languages/csharp/tests/csharp/test_csharp_array_copy_behavior.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// array_copy_behavior
int? maybe = null; int fallback = maybe ?? 26; __Check((fallback == 26).ToString(), "True");
