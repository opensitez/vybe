// vybe-test: csharp/csharp_environment_system/gc_get_total_memory_returns_positive_long
// origin: languages/csharp/tests/csharp/test_csharp_environment_system.rs

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

__P((System.GC.GetTotalMemory(false)>0).ToString());
__Check("True");
