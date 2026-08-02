// vybe-test: csharp/csharp_anonymous_object_basics/anonymous_object_basics_feature_entropy
// origin: languages/csharp/tests/csharp/test_csharp_anonymous_object_basics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// anonymous_object_basics
string feature = "anonymous_object_basics:38"; __Check((feature.Length >= 1).ToString(), "True");
