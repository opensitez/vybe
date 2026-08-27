// vybe-test: csharp/csharp_oop_lock_object_semantics/lock_object_case_11

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

System.Threading.Lock lk = new System.Threading.Lock();
int count = 0;
lock (lk) {
    count += 11;
}
__P(count.ToString());
__Check("11");
