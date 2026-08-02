// vybe-test: csharp/csharp_type_casting/int_to_double_implicit_widens_without_loss
// origin: languages/csharp/tests/csharp/test_csharp_type_casting.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int i = 5; double d = i; __Check((d).ToString(), "5");
