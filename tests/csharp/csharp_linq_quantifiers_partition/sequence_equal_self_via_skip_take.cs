// vybe-test: csharp/csharp_linq_quantifiers_partition/sequence_equal_self_via_skip_take
// origin: languages/csharp/tests/csharp/test_csharp_linq_quantifiers_partition.rs

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

var a=new[]{1,2,3,4};
__P((a.Skip(1).Take(2).SequenceEqual(new[]{2,3})).ToString());
__Check("True");
