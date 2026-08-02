// vybe-test: csharp/csharp_generic_inference_matrix/generic_inference_matrix_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_generic_inference_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// generic_inference_matrix
var map = new System.Collections.Generic.Dictionary<int, int>(); map[81] = 82; __Check((map.ContainsKey(81) && map[81] == 82).ToString(), "True");
