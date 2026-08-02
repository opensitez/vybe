// vybe-test: csharp/csharp_delegates_advanced/predicate_combined_with_and
// origin: languages/csharp/tests/csharp/test_csharp_delegates_advanced.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Predicate<int> positive=x=>x>0;
System.Predicate<int> even=x=>x%2==0;
System.Predicate<int> both=x=>positive(x)&&even(x);
__Check((both(4)).ToString(), "True"); __Check((both(-2)).ToString(), "False"); __Check((both(3)).ToString(), "False");
