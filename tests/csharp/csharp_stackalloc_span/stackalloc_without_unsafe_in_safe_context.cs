// vybe-test: csharp/csharp_stackalloc_span/stackalloc_without_unsafe_in_safe_context
// origin: languages/csharp/tests/csharp/test_csharp_stackalloc_span.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Span<int> nums=stackalloc int[3]{7,8,9}; __Check((nums[2]).ToString(), "9");
