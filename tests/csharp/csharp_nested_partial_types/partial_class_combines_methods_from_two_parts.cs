// vybe-test: csharp/csharp_nested_partial_types/partial_class_combines_methods_from_two_parts
// origin: languages/csharp/tests/csharp/test_csharp_nested_partial_types.rs

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

partial class Worker {
    public string First() { return "one"; }
}
partial class Worker {
    public string Second() { return "two"; }
}
var worker = new Worker();
__P((worker.First()).ToString());
__P((worker.Second()).ToString());
__Check("one\ntwo");
