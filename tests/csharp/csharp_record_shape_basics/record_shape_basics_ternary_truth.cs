// vybe-test: csharp/csharp_record_shape_basics/record_shape_basics_ternary_truth
// origin: languages/csharp/tests/csharp/test_csharp_record_shape_basics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// record_shape_basics
int seed = 39; bool cond = seed % 2 == 0; __Check((cond || !cond).ToString(), "True");
