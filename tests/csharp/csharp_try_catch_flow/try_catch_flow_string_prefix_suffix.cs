// vybe-test: csharp/csharp_try_catch_flow/try_catch_flow_string_prefix_suffix
// origin: languages/csharp/tests/csharp/test_csharp_try_catch_flow.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// try_catch_flow
string feature = "try_catch_flow"; __Check((feature.Substring(0, 1) == feature[0].ToString()).ToString(), "True");
