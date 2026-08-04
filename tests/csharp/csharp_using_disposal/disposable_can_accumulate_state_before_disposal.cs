// vybe-test: csharp/csharp_using_disposal/disposable_can_accumulate_state_before_disposal
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

using System; class Buffer : IDisposable { int count; public void Add() { count++; } public void Dispose() { __P((count).ToString()); } } using (var buffer = new Buffer()) { buffer.Add(); buffer.Add(); }
__Check("2");
