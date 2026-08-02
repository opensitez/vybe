// vybe-test: csharp/csharp_record_shape_basics/record_shape_basics_ordering_pair
// origin: languages/csharp/tests/csharp/test_csharp_record_shape_basics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// record_shape_basics
int seed = 39; int right = seed + 1; __Check((seed < right).ToString(), "True");
