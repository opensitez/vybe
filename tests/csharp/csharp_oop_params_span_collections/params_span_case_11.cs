// vybe-test: csharp/csharp_oop_params_span_collections/params_span_case_11

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

int Sum(params ReadOnlySpan<int> values) {
    int s = 0;
    foreach (var v in values) s += v;
    return s;
}
int res = Sum(11, 12);
__P(res.ToString());
__Check("23");
