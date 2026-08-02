// vybe-test: csharp/csharp_linq_join_matrix/linq_join_matrix_string_first_char
// origin: languages/csharp/tests/csharp/test_csharp_linq_join_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// linq_join_matrix
string feature = "linq_join_matrix"; __Check((feature[0] == feature[0]).ToString(), "True");
