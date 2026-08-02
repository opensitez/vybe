// vybe-test: csharp/csharp_generic_inference_matrix/generic_inference_matrix_hashset_uniqueness
// origin: languages/csharp/tests/csharp/test_csharp_generic_inference_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// generic_inference_matrix
var set = new System.Collections.Generic.HashSet<int>(); set.Add(81); set.Add(81); __Check((set.Count == 1).ToString(), "True");
