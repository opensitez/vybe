// vybe-test: csharp/csharp_linq_default_if_empty_fallback/linq_default_if_empty_case_13

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

var empty = new int[0];
int val = empty.DefaultIfEmpty(1300).First();
__P(val.ToString());
__Check("1300");
