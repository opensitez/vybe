// vybe-test: csharp/csharp_tuple_projection_checks/tuple_projection_checks_string_first_char
// origin: languages/csharp/tests/csharp/test_csharp_tuple_projection_checks.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// tuple_projection_checks
string feature = "tuple_projection_checks"; __Check((feature[0] == feature[0]).ToString(), "True");
