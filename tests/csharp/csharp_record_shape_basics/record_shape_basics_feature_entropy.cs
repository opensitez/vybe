// vybe-test: csharp/csharp_record_shape_basics/record_shape_basics_feature_entropy
// origin: languages/csharp/tests/csharp/test_csharp_record_shape_basics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// record_shape_basics
string feature = "record_shape_basics:39"; __Check((feature.Length >= 1).ToString(), "True");
