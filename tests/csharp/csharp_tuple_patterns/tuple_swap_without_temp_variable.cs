// vybe-test: csharp/csharp_tuple_patterns/tuple_swap_without_temp_variable
// origin: languages/csharp/tests/csharp/test_csharp_tuple_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int x=1,y=2;
(x,y)=(y,x);
__Check((x).ToString(), "2"); __Check((y).ToString(), "1");
