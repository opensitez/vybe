// vybe-test: csharp/csharp_auto_property_defaults/auto_property_defaults_list_count_contract
// origin: languages/csharp/tests/csharp/test_csharp_auto_property_defaults.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// auto_property_defaults
var values = new System.Collections.Generic.List<int> { 65, 66, 65 }; __Check((values.Count == 3).ToString(), "True");
