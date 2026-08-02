// vybe-test: csharp/csharp_threading_primitives/lazy_factory_runs_once_on_first_value_access
// origin: languages/csharp/tests/csharp/test_csharp_threading_primitives.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int calls = 0;
var lazy = new System.Lazy<int>(() => { calls++; return 7; });
__Check((calls).ToString(), "0");
__Check((lazy.Value).ToString(), "7");
__Check((calls).ToString(), "1");
