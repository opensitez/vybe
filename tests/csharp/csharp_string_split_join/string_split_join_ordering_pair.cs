// vybe-test: csharp/csharp_string_split_join/string_split_join_ordering_pair
// origin: languages/csharp/tests/csharp/test_csharp_string_split_join.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// string_split_join
int seed = 21; int right = seed + 1; __Check((seed < right).ToString(), "True");
