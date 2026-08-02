// vybe-test: csharp/csharp_array_copy_behavior/array_copy_behavior_ternary_truth
// origin: languages/csharp/tests/csharp/test_csharp_array_copy_behavior.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// array_copy_behavior
int seed = 26; bool cond = seed % 2 == 0; __Check((cond || !cond).ToString(), "True");
