// vybe-test: csharp/csharp_record_struct_deep/record_struct_return_method
// origin: languages/csharp/tests/csharp/test_csharp_record_struct_deep.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record struct V(int N); V Make()=>new V(7); __Check((Make().N).ToString(), "7");
