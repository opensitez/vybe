// vybe-test: csharp/csharp_struct_features/struct_passed_by_ref_mutates_caller_copy
// origin: languages/csharp/tests/csharp/test_csharp_struct_features.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

struct Counter { public int N; }
void Increment(ref Counter c) { c.N++; }
var c = new Counter { N=5 };
Increment(ref c);
__Check((c.N).ToString(), "6");
