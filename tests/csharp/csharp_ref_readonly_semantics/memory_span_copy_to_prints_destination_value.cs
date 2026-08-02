// vybe-test: csharp/csharp_ref_readonly_semantics/memory_span_copy_to_prints_destination_value
// origin: languages/csharp/tests/csharp/test_csharp_ref_readonly_semantics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var src=new System.Memory<int>(new int[]{7,8}); int[] dst=new int[2]; src.Span.CopyTo(dst); __Check((dst[1]).ToString(), "8");
