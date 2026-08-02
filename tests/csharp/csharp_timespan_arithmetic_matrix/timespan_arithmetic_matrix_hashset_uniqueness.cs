// vybe-test: csharp/csharp_timespan_arithmetic_matrix/timespan_arithmetic_matrix_hashset_uniqueness
// origin: languages/csharp/tests/csharp/test_csharp_timespan_arithmetic_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// timespan_arithmetic_matrix
var set = new System.Collections.Generic.HashSet<int>(); set.Add(95); set.Add(95); __Check((set.Count == 1).ToString(), "True");
