// vybe-test: csharp/csharp_class_indexers/indexer_on_readonly_wrapper_exposes_underlying_element
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

class ReadWrapper {
    readonly int[] data = { 5, 6 };
    public int this[int i] { get { return data[i]; } }
}
__P((new ReadWrapper()[0]).ToString());
__Check("5");
