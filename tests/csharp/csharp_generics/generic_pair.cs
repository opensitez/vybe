// vybe-test: csharp/csharp_generics/generic_pair
// origin: languages/csharp/tests/csharp/test_csharp_generics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class Pair<T1, T2> {
    public T1 First;
    public T2 Second;
    public Pair(T1 a, T2 b) { First = a; Second = b; }
}
var p = new Pair<string, int>("age", 30);
__Check((p.First).ToString(), "age");
__Check((p.Second).ToString(), "30");
