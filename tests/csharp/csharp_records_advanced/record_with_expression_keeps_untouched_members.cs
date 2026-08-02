// vybe-test: csharp/csharp_records_advanced/record_with_expression_keeps_untouched_members
// origin: languages/csharp/tests/csharp/test_csharp_records_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record Item(string Name, int Count); var item = new Item("pen", 2); var changed = item with { Count = 5 }; __Check((changed.Name).ToString(), "pen");
