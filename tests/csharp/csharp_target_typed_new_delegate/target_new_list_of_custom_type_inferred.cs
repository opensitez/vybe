// vybe-test: csharp/csharp_target_typed_new_delegate/target_new_list_of_custom_type_inferred
// origin: languages/csharp/tests/csharp/test_csharp_target_typed_new_delegate.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Item { public string Name = ""; }
System.Collections.Generic.List<Item> items = new();
items.Add(new Item { Name = "tool" });
__Check((items[0].Name).ToString(), "tool");
