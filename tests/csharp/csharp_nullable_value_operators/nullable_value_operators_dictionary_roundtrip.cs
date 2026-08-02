// vybe-test: csharp/csharp_nullable_value_operators/nullable_value_operators_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_nullable_value_operators.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// nullable_value_operators
var map = new System.Collections.Generic.Dictionary<int, int>(); map[57] = 58; __Check((map.ContainsKey(57) && map[57] == 58).ToString(), "True");
