// vybe-test: csharp/csharp_threading_manual_reset_event_slim/manual_reset_event_slim_case_19

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

using var mres = new System.Threading.ManualResetEventSlim(false);
mres.Set();
bool ok = mres.Wait(100);
__P(ok.ToString());
__Check("True");
