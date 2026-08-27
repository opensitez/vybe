// vybe-test: csharp/csharp_linq_try_get_non_enumerated_count/linq_non_enumerated_case_2

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

var list = new System.Collections.Generic.List<int>() { 1, 2, 3 };
bool ok = list.TryGetNonEnumeratedCount(out int count);
__P(ok.ToString());
__P(count.ToString());
__Check("True\n3");
