// vybe-test: csharp/csharp_concurrency_sync/monitor_try_enter_returns_false_when_already_locked
// origin: languages/csharp/tests/csharp/test_csharp_concurrency_sync.rs

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

object obj=new object();
System.Threading.Monitor.Enter(obj);
bool got=System.Threading.Monitor.TryEnter(obj,0);
System.Threading.Monitor.Exit(obj);
__P((got).ToString());
__Check("False");
