// vybe-test: csharp/csharp_class_indexers/indexer_get_can_compute_derived_value_from_state
// origin: languages/csharp/tests/csharp/test_csharp_class_indexers.rs

string __buf = "";

void __P(string s) {
    __buf = __buf + s + "\n";
}

void __Pr(string s) {
    __buf = __buf + s;
}

// The final WriteLine contributes a trailing newline that the expected line
// vector never carried, so BOTH forms are accepted.
void __Check(string want) {
    if (__buf != want && __buf != want + "\n") {
        Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
        throw new Exception("assertion failed");
    }
}

class Scale {
    int factor = 2;
    public int this[int input] { get { return input * factor; } }
}
__P((new Scale()[5]).ToString());
__Check("10");
