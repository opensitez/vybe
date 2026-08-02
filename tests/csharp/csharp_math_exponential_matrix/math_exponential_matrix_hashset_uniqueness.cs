// vybe-test: csharp/csharp_math_exponential_matrix/math_exponential_matrix_hashset_uniqueness
// origin: languages/csharp/tests/csharp/test_csharp_math_exponential_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// math_exponential_matrix
var set = new System.Collections.Generic.HashSet<int>(); set.Add(103); set.Add(103); __Check((set.Count == 1).ToString(), "True");
