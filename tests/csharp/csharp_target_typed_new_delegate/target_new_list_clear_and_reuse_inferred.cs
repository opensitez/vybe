// vybe-test: csharp/csharp_target_typed_new_delegate/target_new_list_clear_and_reuse_inferred
// origin: languages/csharp/tests/csharp/test_csharp_target_typed_new_delegate.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Collections.Generic.List<int> buf = new() { 1, 2 };
buf.Clear();
buf.Add(9);
__Check((buf[0]).ToString(), "9");
