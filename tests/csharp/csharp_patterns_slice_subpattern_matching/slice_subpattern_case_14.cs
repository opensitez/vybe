// vybe-test: csharp/csharp_patterns_slice_subpattern_matching/slice_subpattern_case_14

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

int[] arr = new int[] { 1, 14, 3 };
if (arr is [1, var mid, 3]) {
    __P("True");
    __P(mid.ToString());
}
__Check("True\n14");
