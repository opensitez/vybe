// vybe-test: csharp/csharp_delegates/action_no_params
// origin: languages/csharp/tests/csharp/test_csharp_delegates.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

Action sayHi = () => __Check(("hi").ToString(), "hi");
sayHi();
