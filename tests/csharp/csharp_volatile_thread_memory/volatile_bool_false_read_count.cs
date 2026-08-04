// vybe-test: csharp/csharp_volatile_thread_memory/volatile_bool_false_read_count
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
    public volatile bool Ready = false;
}
var box = new FlagBox();
__P((box.Ready ? 1 : 0).ToString());
__Check("0");
