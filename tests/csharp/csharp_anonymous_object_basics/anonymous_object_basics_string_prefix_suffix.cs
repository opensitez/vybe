// vybe-test: csharp/csharp_anonymous_object_basics/anonymous_object_basics_string_prefix_suffix
// origin: languages/csharp/tests/csharp/test_csharp_anonymous_object_basics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// anonymous_object_basics
string feature = "anonymous_object_basics"; __Check((feature.Substring(0, 1) == feature[0].ToString()).ToString(), "True");
