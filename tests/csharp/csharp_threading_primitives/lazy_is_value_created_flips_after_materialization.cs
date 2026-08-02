// vybe-test: csharp/csharp_threading_primitives/lazy_is_value_created_flips_after_materialization
// origin: languages/csharp/tests/csharp/test_csharp_threading_primitives.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var lazy = new System.Lazy<int>(() => 3);
__Check((lazy.IsValueCreated).ToString(), "False");
__Check((lazy.Value).ToString(), "3");
__Check((lazy.IsValueCreated).ToString(), "True");
