// vybe-test: csharp/csharp_threading_thread_local_storage/thread_local_case_11

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

using var tl = new System.Threading.ThreadLocal<int>(() => 11);
__P(tl.Value.ToString());
__Check("11");
