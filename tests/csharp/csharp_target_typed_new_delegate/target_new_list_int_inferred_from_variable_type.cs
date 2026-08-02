// vybe-test: csharp/csharp_target_typed_new_delegate/target_new_list_int_inferred_from_variable_type
// origin: languages/csharp/tests/csharp/test_csharp_target_typed_new_delegate.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Collections.Generic.List<int> values = new();
values.Add(7);
__Check((values[0]).ToString(), "7");
