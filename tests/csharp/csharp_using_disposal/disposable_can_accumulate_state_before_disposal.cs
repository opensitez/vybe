// vybe-test: csharp/csharp_using_disposal/disposable_can_accumulate_state_before_disposal
// origin: languages/csharp/tests/csharp/test_csharp_using_disposal.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System; class Buffer : IDisposable { int count; public void Add() { count++; } public void Dispose() { __Check((count).ToString(), "2"); } } using (var buffer = new Buffer()) { buffer.Add(); buffer.Add(); }
