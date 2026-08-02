// vybe-test: csharp/csharp_char_predicate_apis/char_predicate_apis_ternary_truth
// origin: languages/csharp/tests/csharp/test_csharp_char_predicate_apis.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// char_predicate_apis
int seed = 23; bool cond = seed % 2 == 0; __Check((cond || !cond).ToString(), "True");
