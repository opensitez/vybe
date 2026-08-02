// vybe-test: csharp/basics/prefix_decrement
// origin: languages/csharp/tests/csharp/test_basics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var x = 10;
        --x;
        --x;
        __Check((x).ToString(), "8");
