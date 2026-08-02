// vybe-test: csharp/csharp_anonymous_object_basics/anonymous_object_basics_double_identity
// origin: languages/csharp/tests/csharp/test_csharp_anonymous_object_basics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// anonymous_object_basics
double seed = 38; __Check(((seed + 0.5 - 0.5) == seed).ToString(), "True");
