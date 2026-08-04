// vybe-test: csharp/csharp_volatile_thread_memory/volatile_two_instances_isolated_counts
// origin: languages/csharp/tests/csharp/test_csharp_volatile_thread_memory.rs

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

class FlagBox {
    public volatile int Value = 0;
}
var a = new FlagBox();
var b = new FlagBox();
a.Value = 4;
b.Value = 5;
__P((a.Value + b.Value).ToString());
__Check("9");
