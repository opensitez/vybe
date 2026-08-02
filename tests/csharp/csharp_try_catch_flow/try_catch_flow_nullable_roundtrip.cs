// vybe-test: csharp/csharp_try_catch_flow/try_catch_flow_nullable_roundtrip
// origin: languages/csharp/tests/csharp/test_csharp_try_catch_flow.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

// try_catch_flow
int? maybe = 51; __Check((maybe.HasValue && maybe.Value == 51).ToString(), "True");
