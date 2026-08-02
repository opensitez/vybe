// vybe-test: csharp/basics/multi_var_three_variables
// origin: languages/csharp/tests/csharp/test_basics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int x = 10, y = 20, z = 30;
        __Check((x + y + z).ToString(), "60");
