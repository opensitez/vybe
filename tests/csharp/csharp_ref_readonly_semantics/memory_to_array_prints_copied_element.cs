// vybe-test: csharp/csharp_ref_readonly_semantics/memory_to_array_prints_copied_element
// origin: languages/csharp/tests/csharp/test_csharp_ref_readonly_semantics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var memory=new System.Memory<int>(new int[]{9,8,7}); __Check((memory.ToArray()[2]).ToString(), "7");
