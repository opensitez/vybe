// vybe-test: csharp/csharp_extension_methods/extension_method_on_custom_class_can_read_field
// origin: languages/csharp/tests/csharp/test_csharp_extension_methods.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using Demo; class Item { public int Count = 5; } namespace Demo { public static class ItemExt { public static int Double(this Item item) { return item.Count * 2; } } } __Check((new Item().Double()).ToString(), "10");
