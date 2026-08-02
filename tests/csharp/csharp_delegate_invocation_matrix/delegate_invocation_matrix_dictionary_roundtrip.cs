// vybe-test: csharp/csharp_delegate_invocation_matrix/delegate_invocation_matrix_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_delegate_invocation_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// delegate_invocation_matrix
var map = new System.Collections.Generic.Dictionary<int, int>(); map[74] = 75; __Check((map.ContainsKey(74) && map[74] == 75).ToString(), "True");
