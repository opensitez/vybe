// vybe-test: csharp/csharp_ref_readonly_semantics/ref_readonly_extension_method_on_span
// origin: languages/csharp/tests/csharp/test_csharp_ref_readonly_semantics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

static class SpanExt{public static int First(ref readonly System.Span<int> span)=>span.Length>0?span[0]:-1;} System.Span<int> s=stackalloc int[2]{5,6}; __Check((SpanExt.First(ref s)).ToString(), "5");
