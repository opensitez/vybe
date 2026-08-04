// vybe-test: csharp/csharp_linq_set_ops/union_merges_two_sequences_without_duplicates
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

var result = new[]{1,2,3}.Union(new[]{3,4,5}).OrderBy(x=>x);
foreach(var x in result) __P((x).ToString());
__Check("1\n2\n3\n4\n5");
