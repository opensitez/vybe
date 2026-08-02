// vybe-test: csharp/csharp_guid_parse_matrix/guid_parse_matrix_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_guid_parse_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// guid_parse_matrix
var map = new System.Collections.Generic.Dictionary<int, int>(); map[97] = 98; __Check((map.ContainsKey(97) && map[97] == 98).ToString(), "True");
