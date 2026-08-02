// vybe-test: csharp/csharp_ref_readonly_semantics/memory_span_sequence_equal_prints_false_for_different_values
// origin: languages/csharp/tests/csharp/test_csharp_ref_readonly_semantics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var left=new System.Memory<int>(new int[]{1,2}); var right=new System.Memory<int>(new int[]{1,9}); __Check((left.Span.SequenceEqual(right.Span)).ToString(), "False");
