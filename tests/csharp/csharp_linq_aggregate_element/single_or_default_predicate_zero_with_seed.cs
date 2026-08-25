// vybe-test: csharp/csharp_linq_aggregate_element/single_or_default_predicate_zero_with_seed
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

// `SingleOrDefault(predicate, defaultValue)`, not the reverse — the other
// order is `CS1660: Cannot convert lambda expression to type 'int'`.
__P((new[]{1,2,3}.SingleOrDefault(x=>x>10, 55)).ToString());
__Check("55");
