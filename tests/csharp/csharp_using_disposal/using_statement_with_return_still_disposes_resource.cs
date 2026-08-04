// vybe-test: csharp/csharp_using_disposal/using_statement_with_return_still_disposes_resource
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

using System; class Resource : IDisposable { public void Dispose() { __P(("disposed").ToString()); } } int Read() { using (var resource = new Resource()) { __P(("inside").ToString()); return 5; } } __P((Read()).ToString());
__Check("inside\ndisposed\n5");
