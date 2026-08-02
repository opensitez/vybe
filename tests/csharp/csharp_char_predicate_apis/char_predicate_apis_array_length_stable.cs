// vybe-test: csharp/csharp_char_predicate_apis/char_predicate_apis_array_length_stable
// origin: languages/csharp/tests/csharp/test_csharp_char_predicate_apis.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// char_predicate_apis
int seed = 23; int[] numbers = new int[] { seed, seed + 1, seed + 2 }; __Check((numbers.Length == 3).ToString(), "True");
