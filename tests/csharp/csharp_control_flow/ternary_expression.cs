// vybe-test: csharp/csharp_control_flow/ternary_expression
// origin: languages/csharp/tests/csharp/test_csharp_control_flow.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int x = 5;
string result = x > 3 ? "big" : "small";
__Check((result).ToString(), "big");
