// vybe-test: csharp/control_flow/try_catch_basic
// origin: languages/csharp/tests/csharp/test_control_flow.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

try {
            throw new Exception("oops");
        } catch (Exception e) {
            __Check(("caught").ToString(), "caught");
        }
