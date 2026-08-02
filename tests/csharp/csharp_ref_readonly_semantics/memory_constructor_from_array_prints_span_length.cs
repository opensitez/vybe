// vybe-test: csharp/csharp_ref_readonly_semantics/memory_constructor_from_array_prints_span_length
// origin: languages/csharp/tests/csharp/test_csharp_ref_readonly_semantics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var memory=new System.Memory<int>(new int[]{1,2,3}); __Check((memory.Length).ToString(), "3");
