// vybe-test: csharp/csharp_hashset_membership_matrix/hashset_membership_matrix_nullable_fallback
// origin: languages/csharp/tests/csharp/test_csharp_hashset_membership_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// hashset_membership_matrix
int? maybe = null; int fallback = maybe ?? 33; __Check((fallback == 33).ToString(), "True");
