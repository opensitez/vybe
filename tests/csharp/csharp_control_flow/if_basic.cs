// vybe-test: csharp/csharp_control_flow/if_basic
// origin: languages/csharp/tests/csharp/test_csharp_control_flow.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int x = 5;
if (x > 3) {
    __Check(("big").ToString(), "big");
}
