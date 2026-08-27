// vybe-test: csharp/csharp_threading_async_local_context_flow/async_local_case_11

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

var al = new System.Threading.AsyncLocal<string>();
al.Value = "Context_11";
__P(al.Value);
__Check("Context_11");
