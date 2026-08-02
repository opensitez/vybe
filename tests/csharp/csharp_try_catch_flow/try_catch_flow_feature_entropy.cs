// vybe-test: csharp/csharp_try_catch_flow/try_catch_flow_feature_entropy
// origin: languages/csharp/tests/csharp/test_csharp_try_catch_flow.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// try_catch_flow
string feature = "try_catch_flow:51"; __Check((feature.Length >= 1).ToString(), "True");
