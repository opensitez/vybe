// vybe-test: csharp/control_flow/switch_basic
// origin: languages/csharp/tests/csharp/test_control_flow.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var x = 2;
        var result = "";
        switch (x) {
            case 1: result = "one"; break;
            case 2: result = "two"; break;
            case 3: result = "three"; break;
            default: result = "other"; break;
        }
        __Check((result).ToString(), "two");
