// vybe-test: csharp/csharp_ref_readonly_semantics/memory_slice_offset_and_length_prints_inner_value
// origin: languages/csharp/tests/csharp/test_csharp_ref_readonly_semantics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var memory=new System.Memory<int>(new int[]{2,4,6,8}); __Check((memory.Slice(1,2).Span[1]).ToString(), "6");
