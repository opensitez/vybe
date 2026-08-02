// vybe-test: csharp/csharp_threading_primitives/weak_reference_target_is_alive_while_strong_reference_exists
// origin: languages/csharp/tests/csharp/test_csharp_threading_primitives.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var strong = new object();
var weak = new System.WeakReference(strong);
__Check((weak.IsAlive).ToString(), "True");
