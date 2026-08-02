// vybe-test: csharp/csharp_ref_readonly_semantics/memory_span_try_copy_to_reports_success
// origin: languages/csharp/tests/csharp/test_csharp_ref_readonly_semantics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var src=new System.Memory<int>(new int[]{1,2}); int[] dst=new int[2]; __Check((src.Span.TryCopyTo(dst)).ToString(), "True");
