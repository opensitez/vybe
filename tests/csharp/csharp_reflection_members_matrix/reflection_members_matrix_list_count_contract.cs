// vybe-test: csharp/csharp_reflection_members_matrix/reflection_members_matrix_list_count_contract
// origin: languages/csharp/tests/csharp/test_csharp_reflection_members_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// reflection_members_matrix
var values = new System.Collections.Generic.List<int> { 92, 93, 92 }; __Check((values.Count == 3).ToString(), "True");
