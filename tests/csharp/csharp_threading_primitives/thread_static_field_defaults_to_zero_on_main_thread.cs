// vybe-test: csharp/csharp_threading_primitives/thread_static_field_defaults_to_zero_on_main_thread
// origin: languages/csharp/tests/csharp/test_csharp_threading_primitives.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Counter {
    [System.ThreadStatic]
    public static int Value;
}
__Check((Counter.Value).ToString(), "0");
