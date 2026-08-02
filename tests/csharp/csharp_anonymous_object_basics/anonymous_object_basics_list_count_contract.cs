// vybe-test: csharp/csharp_anonymous_object_basics/anonymous_object_basics_list_count_contract
// origin: languages/csharp/tests/csharp/test_csharp_anonymous_object_basics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// anonymous_object_basics
var values = new System.Collections.Generic.List<int> { 38, 39, 38 }; __Check((values.Count == 3).ToString(), "True");
