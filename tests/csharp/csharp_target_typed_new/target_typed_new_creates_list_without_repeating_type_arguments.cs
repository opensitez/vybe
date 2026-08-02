// vybe-test: csharp/csharp_target_typed_new/target_typed_new_creates_list_without_repeating_type_arguments
// origin: languages/csharp/tests/csharp/test_csharp_target_typed_new.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Collections.Generic.List<int> values = new();
values.Add(7);
__Check((values[0]).ToString(), "7");
