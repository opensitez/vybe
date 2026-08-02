// vybe-test: csharp/csharp_extension_methods/extension_method_on_list_can_report_item_count
// origin: languages/csharp/tests/csharp/test_csharp_extension_methods.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

using Demo; using System.Collections.Generic; namespace Demo { public static class ListExt { public static string Describe<T>(this List<T> values) { return "count=" + values.Count; } } } __Check((new List<int> { 1, 2 }.Describe()).ToString(), "count=2");
