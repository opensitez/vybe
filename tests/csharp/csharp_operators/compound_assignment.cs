// vybe-test: csharp/csharp_operators/compound_assignment
// origin: languages/csharp/tests/csharp/test_csharp_operators.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int x = 10;
x += 5; __Check((x).ToString(), "15");
x -= 3; __Check((x).ToString(), "12");
x *= 2; __Check((x).ToString(), "24");
x /= 4; __Check((x).ToString(), "6");
x %= 5; __Check((x).ToString(), "1");
