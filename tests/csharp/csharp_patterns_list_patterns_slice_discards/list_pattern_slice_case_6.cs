// vybe-test: csharp/csharp_patterns_list_patterns_slice_discards/list_pattern_slice_case_6

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

int[] arr = new int[] { 6, 2, 3, 60 };
bool matches = arr is [6, .., 60];
__P(matches.ToString());
__Check("True");
