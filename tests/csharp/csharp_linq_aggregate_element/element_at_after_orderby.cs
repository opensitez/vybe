// vybe-test: csharp/csharp_linq_aggregate_element/element_at_after_orderby
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

__P((new[]{3,1,2}.OrderBy(x=>x).ElementAt(1)).ToString());
__Check("2");
