// vybe-test: csharp/csharp_using_disposal/using_block_can_allocate_and_return_computed_value
// origin: languages/csharp/tests/csharp/test_csharp_using_disposal.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using System; class Resource : IDisposable { public int Value => 4; public void Dispose() { __Check(("disposed").ToString(), "8"); } } using (var resource = new Resource()) { __Check((resource.Value * 2).ToString(), "disposed"); }
