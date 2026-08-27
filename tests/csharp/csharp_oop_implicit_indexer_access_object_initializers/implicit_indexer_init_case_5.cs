// vybe-test: csharp/csharp_oop_implicit_indexer_access_object_initializers/implicit_indexer_init_case_5

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

var list = new int[] { 10, 20, 30 };
int last = list[^1];
__P(last.ToString());
__Check("30");
