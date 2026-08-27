// vybe-test: csharp/csharp_threading_periodic_timer_tick_flow/periodic_timer_case_12

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

using var timer = new System.Threading.PeriodicTimer(TimeSpan.FromMilliseconds(100));
__P((timer != null).ToString());
__Check("True");
