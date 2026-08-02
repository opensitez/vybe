// vybe-test: csharp/csharp_list_operations/to_array_converts_list_to_fixed_array
// origin: languages/csharp/tests/csharp/test_csharp_list_operations.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var list = new System.Collections.Generic.List<int>{7,8,9};
var arr = list.ToArray();
__Check((arr.GetType().IsArray).ToString(), "True");
