// vybe-test: csharp/csharp_try_catch_flow/try_catch_flow_double_identity
// origin: languages/csharp/tests/csharp/test_csharp_try_catch_flow.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// try_catch_flow
double seed = 51; __Check(((seed + 0.5 - 0.5) == seed).ToString(), "True");
