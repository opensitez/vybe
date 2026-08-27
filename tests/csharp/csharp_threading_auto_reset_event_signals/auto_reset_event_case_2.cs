// vybe-test: csharp/csharp_threading_auto_reset_event_signals/auto_reset_event_case_2

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

using var are = new System.Threading.AutoResetEvent(true);
bool w1 = are.WaitOne(100);
bool w2 = are.WaitOne(10);
__P(w1.ToString());
__P(w2.ToString());
__Check("True\nFalse");
