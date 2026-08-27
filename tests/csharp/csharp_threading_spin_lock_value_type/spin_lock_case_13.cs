// vybe-test: csharp/csharp_threading_spin_lock_value_type/spin_lock_case_13

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

var sl = new System.Threading.SpinLock();
bool lockTaken = false;
sl.Enter(ref lockTaken);
__P(lockTaken.ToString());
sl.Exit();
__Check("True");
