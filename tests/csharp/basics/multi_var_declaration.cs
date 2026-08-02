// vybe-test: csharp/basics/multi_var_declaration
// origin: languages/csharp/tests/csharp/test_basics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int a = 1, b = 2;
        __Check((a + b).ToString(), "3");
