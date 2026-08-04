// vybe-test: csharp/csharp_structs_value_semantics/passing_struct_by_ref_allows_mutation_of_original
// origin: languages/csharp/tests/csharp/test_csharp_structs_value_semantics.rs

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

struct Counter { public int Value; } void Bump(ref Counter counter) { counter.Value++; } var counter = new Counter { Value = 2 }; Bump(ref counter); __P((counter.Value).ToString());
__Check("3");
