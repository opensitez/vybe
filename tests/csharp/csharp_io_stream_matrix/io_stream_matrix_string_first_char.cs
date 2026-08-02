// vybe-test: csharp/csharp_io_stream_matrix/io_stream_matrix_string_first_char
// origin: languages/csharp/tests/csharp/test_csharp_io_stream_matrix.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// io_stream_matrix
string feature = "io_stream_matrix"; __Check((feature[0] == feature[0]).ToString(), "True");
