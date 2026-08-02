// vybe-test: csharp/csharp_using_declarations/using_var_memory_stream_length_ok_before_scope_end
// origin: languages/csharp/tests/csharp/test_csharp_using_declarations.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using var ms=new System.IO.MemoryStream(new byte[]{1,2,3}); __Check((ms.Length).ToString(), "3");
