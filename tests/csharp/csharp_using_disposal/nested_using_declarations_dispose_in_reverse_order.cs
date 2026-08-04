// vybe-test: csharp/csharp_using_disposal/nested_using_declarations_dispose_in_reverse_order
// origin: languages/csharp/tests/csharp/test_csharp_using_disposal.rs

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

using System; class Resource : IDisposable { string name; public Resource(string name) { this.name = name; } public void Dispose() { __P((name).ToString()); } } using var first = new Resource("first"); using var second = new Resource("second"); __P(("done").ToString());
__Check("done\nsecond\nfirst");
