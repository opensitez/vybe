// vybe-test: csharp/csharp_casting_patterns/direct_cast_succeeds_for_compatible_type
// origin: languages/csharp/tests/csharp/test_csharp_casting_patterns.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

object o=3.14;
double d=(double)o;
__Check((d).ToString(), "3.14");
