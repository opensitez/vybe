// vybe-test: csharp/csharp_static_classes/static_field_shared_across_all_callers
// origin: languages/csharp/tests/csharp/test_csharp_static_classes.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

static class Counter { public static int Count = 0; }
Counter.Count++;
Counter.Count++;
__Check((Counter.Count).ToString(), "2");
