// vybe-test: csharp/csharp_linq_aggregate_element/single_or_default_predicate_many_with_seed
// origin: languages/csharp/tests/csharp/test_csharp_linq_aggregate_element.rs

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

// Signature is `SingleOrDefault(predicate, defaultValue)` — the other order
// does not compile (`CS1660`). Two elements match, so this throws.
try { __P((new[]{2,2,3}.SingleOrDefault(x=>x==2, 77)).ToString()); }
catch (InvalidOperationException) { __P("threw"); }
__Check("threw");
