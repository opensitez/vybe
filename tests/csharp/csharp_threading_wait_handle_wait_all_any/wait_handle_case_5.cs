// vybe-test: csharp/csharp_threading_wait_handle_wait_all_any/wait_handle_case_5

string __buf = "";

void __P(string s) {
    __buf = __buf + s + "\n";
}

void __Pr(string s) {
    __buf = __buf + s;
}

void __Check(string want) {
    if (__buf != want && __buf != want + "\n") {
        Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
        throw new Exception("assertion failed");
    }
}

using var e1 = new System.Threading.ManualResetEvent(true);
using var e2 = new System.Threading.ManualResetEvent(true);
bool ok = System.Threading.WaitHandle.WaitAll(new System.Threading.WaitHandle[] { e1, e2 }, 100);
__P(ok.ToString());
__Check("True");
