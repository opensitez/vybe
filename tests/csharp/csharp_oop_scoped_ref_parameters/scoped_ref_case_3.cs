// vybe-test: csharp/csharp_oop_scoped_ref_parameters/scoped_ref_case_3

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

void Process(scoped ReadOnlySpan<int> span) {
    __P(span.Length.ToString());
}
Process(new int[] { 3, 4 });
__Check("2");
