// vybe-test: csharp/csharp_static_abstract_interfaces/static_abstract_decimal_from_string
// origin: languages/csharp/tests/csharp/test_csharp_static_abstract_interfaces.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface IDec<T> where T:IDec<T>{static abstract T Parse(string s);}
struct Money:IDec<Money>{public decimal Amount; public static Money Parse(string s)=>new Money{Amount=decimal.Parse(s)};}
__Check((Money.Parse("12.5").Amount).ToString(), "12.5");
