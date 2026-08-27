// vybe-test: csharp/csharp_threading_spin_wait_once/spin_wait_case_19

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

var sw = new System.Threading.SpinWait();
sw.SpinOnce();
__P(sw.Count.ToString());
__Check("1");
