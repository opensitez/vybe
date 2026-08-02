// vybe-test: csharp/csharp_checked_unchecked/unchecked_block_wraps_silently_on_int_overflow
// origin: languages/csharp/tests/csharp/test_csharp_checked_unchecked.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

unchecked{int x=int.MaxValue; x++; __Check((x==int.MinValue).ToString(), "True");}
