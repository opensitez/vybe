// vybe-test: csharp/csharp_value_ref_semantics/passing_class_by_reference_mutates_caller
// origin: languages/csharp/tests/csharp/test_csharp_value_ref_semantics.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

class C{public int V;}
void Mutate(C c){c.V=999;}
var c=new C{V=1};
Mutate(c);
__Check((c.V).ToString(), "999");
