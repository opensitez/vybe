// vybe-test: csharp/csharp_linq_quantifiers_partition/all_and_any_combined_count
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

var data=new[]{2,4,6,8};
__P((data.All(x=>x%2==0)?1:0).ToString());
__P((data.Any(x=>x>5)?1:0).ToString());
__Check("1\n1");
