// vybe-test: csharp/csharp_threading_barrier_phase_sync/barrier_case_1

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

using var barrier = new System.Threading.Barrier(1);
barrier.SignalAndWait(100);
__P(barrier.CurrentPhaseNumber.ToString());
__Check("1");
