// vybe-test: csharp/csharp_threading_primitives/weak_reference_try_get_target_returns_false_after_target_collected
// origin: languages/csharp/tests/csharp/test_csharp_threading_primitives.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.WeakReference weak;
void Create() {
    weak = new System.WeakReference(new object());
}
Create();
System.GC.Collect();
System.GC.WaitForPendingFinalizers();
__Check((weak.TryGetTarget(out var target)).ToString(), "False");
