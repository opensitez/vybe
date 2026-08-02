// vybe-test: csharp/csharp_timespan_arithmetic_matrix/timespan_arithmetic_matrix_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_timespan_arithmetic_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// timespan_arithmetic_matrix
var map = new System.Collections.Generic.Dictionary<int, int>(); map[95] = 96; __Check((map.ContainsKey(95) && map[95] == 96).ToString(), "True");
