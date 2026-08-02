// vybe-test: csharp/csharp_struct_features/passing_struct_to_method_copies_value
// origin: languages/csharp/tests/csharp/test_csharp_struct_features.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

struct Counter { public int N; }
void Increment(Counter c) { c.N++; }
var c = new Counter { N=5 };
Increment(c);
__Check((c.N).ToString(), "5");
