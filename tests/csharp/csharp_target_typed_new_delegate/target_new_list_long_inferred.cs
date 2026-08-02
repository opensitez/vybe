// vybe-test: csharp/csharp_target_typed_new_delegate/target_new_list_long_inferred
// origin: languages/csharp/tests/csharp/test_csharp_target_typed_new_delegate.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Collections.Generic.List<long> longs = new();
longs.Add(9000000000L);
__Check((longs[0]).ToString(), "9000000000");
