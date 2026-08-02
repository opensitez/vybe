// vybe-test: csharp/csharp_ref_readonly_semantics/memory_span_slice_prints_middle_element
// origin: languages/csharp/tests/csharp/test_csharp_ref_readonly_semantics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var memory=new System.Memory<int>(new int[]{10,20,30,40}); __Check((memory.Span.Slice(1,2)[1]).ToString(), "30");
