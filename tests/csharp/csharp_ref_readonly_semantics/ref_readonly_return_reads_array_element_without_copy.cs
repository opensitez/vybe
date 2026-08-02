// vybe-test: csharp/csharp_ref_readonly_semantics/ref_readonly_return_reads_array_element_without_copy
// origin: languages/csharp/tests/csharp/test_csharp_ref_readonly_semantics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] data={10,20,30}; ref readonly int Peek(int i)=>ref data[i]; __Check((Peek(1)).ToString(), "20");
