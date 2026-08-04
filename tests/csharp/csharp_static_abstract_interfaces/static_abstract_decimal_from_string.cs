// vybe-test: csharp/csharp_static_abstract_interfaces/static_abstract_decimal_from_string
// origin: languages/csharp/tests/csharp/test_csharp_static_abstract_interfaces.rs

string __buf = "";

void __P(string s) {
    __buf = __buf + s + "\n";
}

void __Pr(string s) {
    __buf = __buf + s;
}

// The final WriteLine contributes a trailing newline that the expected line
// vector never carried, so BOTH forms are accepted.
void __Check(string want) {
    if (__buf != want && __buf != want + "\n") {
        Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
        throw new Exception("assertion failed");
    }
}

interface IDec<T> where T:IDec<T>{static abstract T Parse(string s);}
struct Money:IDec<Money>{public decimal Amount; public static Money Parse(string s)=>new Money{Amount=decimal.Parse(s)};}
__P((Money.Parse("12.5").Amount).ToString());
__Check("12.5");
