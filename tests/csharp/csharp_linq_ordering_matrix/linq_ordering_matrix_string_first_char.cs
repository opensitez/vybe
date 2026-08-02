// vybe-test: csharp/csharp_linq_ordering_matrix/linq_ordering_matrix_string_first_char
// origin: languages/csharp/tests/csharp/test_csharp_linq_ordering_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// linq_ordering_matrix
string feature = "linq_ordering_matrix"; __Check((feature[0] == feature[0]).ToString(), "True");
