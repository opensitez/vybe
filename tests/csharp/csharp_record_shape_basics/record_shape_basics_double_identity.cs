// vybe-test: csharp/csharp_record_shape_basics/record_shape_basics_double_identity
// origin: languages/csharp/tests/csharp/test_csharp_record_shape_basics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// record_shape_basics
double seed = 39; __Check(((seed + 0.5 - 0.5) == seed).ToString(), "True");
