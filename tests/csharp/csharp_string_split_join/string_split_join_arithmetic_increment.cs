// vybe-test: csharp/csharp_string_split_join/string_split_join_arithmetic_increment
// origin: languages/csharp/tests/csharp/test_csharp_string_split_join.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// string_split_join
int seed = 21; __Check((seed + 1 > seed).ToString(), "True");
