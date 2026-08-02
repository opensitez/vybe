// vybe-test: csharp/csharp_try_catch_flow/try_catch_flow_ordering_pair
// origin: languages/csharp/tests/csharp/test_csharp_try_catch_flow.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// try_catch_flow
int seed = 51; int right = seed + 1; __Check((seed < right).ToString(), "True");
