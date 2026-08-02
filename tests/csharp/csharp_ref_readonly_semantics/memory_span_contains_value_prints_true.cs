// vybe-test: csharp/csharp_ref_readonly_semantics/memory_span_contains_value_prints_true
// origin: languages/csharp/tests/csharp/test_csharp_ref_readonly_semantics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var memory=new System.Memory<int>(new int[]{2,4,6}); __Check((memory.Span.Contains(4)).ToString(), "True");
