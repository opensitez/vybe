// vybe-test: csharp/csharp_loops/nested_foreach_produces_cartesian_pair_count
// origin: languages/csharp/tests/csharp/test_csharp_loops.rs

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

int count=0;
foreach(var a in new[]{1,2})
    foreach(var b in new[]{1,2,3})
        count++;
__P((count).ToString());
__Check("6");
