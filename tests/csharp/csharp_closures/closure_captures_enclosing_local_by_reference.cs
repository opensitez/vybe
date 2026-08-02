// vybe-test: csharp/csharp_closures/closure_captures_enclosing_local_by_reference
// origin: languages/csharp/tests/csharp/test_csharp_closures.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int x = 1;
System.Action inc = () => x++;
inc(); inc();
__Check((x).ToString(), "3");
