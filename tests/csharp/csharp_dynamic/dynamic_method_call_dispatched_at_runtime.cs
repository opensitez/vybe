// vybe-test: csharp/csharp_dynamic/dynamic_method_call_dispatched_at_runtime
// origin: languages/csharp/tests/csharp/test_csharp_dynamic.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

object o="hello";
dynamic d=o;
__Check((d.ToUpper()).ToString(), "HELLO");
