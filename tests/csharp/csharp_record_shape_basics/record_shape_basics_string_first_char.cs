// vybe-test: csharp/csharp_record_shape_basics/record_shape_basics_string_first_char
// origin: languages/csharp/tests/csharp/test_csharp_record_shape_basics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// record_shape_basics
string feature = "record_shape_basics"; __Check((feature[0] == feature[0]).ToString(), "True");
