// vybe-test: csharp/csharp_threading_countdown_event_barrier/countdown_event_case_12

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

using var cde = new System.Threading.CountdownEvent(2);
cde.Signal();
cde.Signal();
bool ok = cde.Wait(100);
__P(ok.ToString());
__Check("True");
