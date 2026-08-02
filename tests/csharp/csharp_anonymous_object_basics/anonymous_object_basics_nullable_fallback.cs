// vybe-test: csharp/csharp_anonymous_object_basics/anonymous_object_basics_nullable_fallback
// origin: languages/csharp/tests/csharp/test_csharp_anonymous_object_basics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// anonymous_object_basics
int? maybe = null; int fallback = maybe ?? 38; __Check((fallback == 38).ToString(), "True");
