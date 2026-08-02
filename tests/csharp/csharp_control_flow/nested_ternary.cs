// vybe-test: csharp/csharp_control_flow/nested_ternary
// origin: languages/csharp/tests/csharp/test_csharp_control_flow.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int x = 5;
string r = x > 10 ? "big" : x > 3 ? "medium" : "small";
__Check((r).ToString(), "medium");
