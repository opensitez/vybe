// vybe-test: csharp/csharp_reflection_members_matrix/reflection_members_matrix_dictionary_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_reflection_members_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// reflection_members_matrix
var map = new System.Collections.Generic.Dictionary<int, int>(); map[92] = 93; __Check((map.ContainsKey(92) && map[92] == 93).ToString(), "True");
