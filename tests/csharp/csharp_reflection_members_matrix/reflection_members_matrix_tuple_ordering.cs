// vybe-test: csharp/csharp_reflection_members_matrix/reflection_members_matrix_tuple_ordering
// origin: languages/csharp/tests/csharp/test_csharp_reflection_members_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// reflection_members_matrix
var tuple = (left: 92, right: 93); __Check((tuple.left < tuple.right).ToString(), "True");
