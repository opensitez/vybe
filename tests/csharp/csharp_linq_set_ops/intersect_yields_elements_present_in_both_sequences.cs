// vybe-test: csharp/csharp_linq_set_ops/intersect_yields_elements_present_in_both_sequences
// origin: languages/csharp/tests/csharp/test_csharp_linq_set_ops.rs

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

var result = new[]{1,2,3,4}.Intersect(new[]{2,4,6}).OrderBy(x=>x);
foreach(var x in result) __P((x).ToString());
__Check("2\n4");
