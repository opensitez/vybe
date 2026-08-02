// vybe-test: csharp/csharp_ref_readonly_semantics/memory_span_fill_sets_uniform_value
// origin: languages/csharp/tests/csharp/test_csharp_ref_readonly_semantics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var memory=new System.Memory<int>(new int[]{0,0,0}); memory.Span.Fill(4); __Check((memory.Span[2]).ToString(), "4");
