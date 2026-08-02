// vybe-test: csharp/csharp_closures/two_closures_share_same_captured_variable
// origin: languages/csharp/tests/csharp/test_csharp_closures.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int shared = 0;
System.Action add = () => shared++;
System.Func<int> read = () => shared;
add(); add();
__Check((read()).ToString(), "2");
