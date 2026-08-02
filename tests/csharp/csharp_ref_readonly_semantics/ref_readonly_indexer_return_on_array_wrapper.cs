// vybe-test: csharp/csharp_ref_readonly_semantics/ref_readonly_indexer_return_on_array_wrapper
// origin: languages/csharp/tests/csharp/test_csharp_ref_readonly_semantics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Buffer{private int[] _data={5,6,7}; public ref readonly int this[int i]=>ref _data[i];} var b=new Buffer(); __Check((b[2]).ToString(), "7");
