// vybe-test: csharp/csharp_threading_primitives/weak_reference_try_get_target_returns_false_after_target_collected
// origin: languages/csharp/tests/csharp/test_csharp_threading_primitives.rs

string __buf = "";

void __P(string s) {
    __buf = __buf + s + "\n";
}

void __Pr(string s) {
    __buf = __buf + s;
}

// The final WriteLine contributes a trailing newline that the expected line
// vector never carried, so BOTH forms are accepted.
void __Check(string want) {
    if (__buf != want && __buf != want + "\n") {
        Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
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
__P((weak.TryGetTarget(out var target)).ToString());
__Check("False");
