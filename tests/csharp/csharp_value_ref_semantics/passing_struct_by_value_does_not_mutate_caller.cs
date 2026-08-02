// vybe-test: csharp/csharp_value_ref_semantics/passing_struct_by_value_does_not_mutate_caller
// origin: languages/csharp/tests/csharp/test_csharp_value_ref_semantics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

struct S{public int V;}
void Mutate(S s){s.V=999;}
var s=new S{V=1};
Mutate(s);
__Check((s.V).ToString(), "1");
