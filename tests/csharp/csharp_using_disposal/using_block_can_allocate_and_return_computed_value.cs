// vybe-test: csharp/csharp_using_disposal/using_block_can_allocate_and_return_computed_value
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

using System; class Resource : IDisposable { public int Value => 4; public void Dispose() { __P(("disposed").ToString()); } } using (var resource = new Resource()) { __P((resource.Value * 2).ToString()); }
__Check("8\ndisposed");
