// vybe-test: csharp/csharp_ref_readonly_semantics/memory_length_matches_array_after_slice
// origin: languages/csharp/tests/csharp/test_csharp_ref_readonly_semantics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var memory=new System.Memory<int>(new int[]{1,2,3,4,5}); __Check((memory.Slice(2).Length).ToString(), "3");
