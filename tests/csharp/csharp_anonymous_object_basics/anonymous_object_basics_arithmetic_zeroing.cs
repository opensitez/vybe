// vybe-test: csharp/csharp_anonymous_object_basics/anonymous_object_basics_arithmetic_zeroing
// origin: languages/csharp/tests/csharp/test_csharp_anonymous_object_basics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// anonymous_object_basics
int seed = 38; __Check((seed - seed == 0).ToString(), "True");
