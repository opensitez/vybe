// vybe-test: csharp/csharp_linq_set_ops/distinct_removes_duplicate_elements
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

var result = new[]{1,2,2,3,1}.Distinct().OrderBy(x=>x);
foreach(var x in result) __P((x).ToString());
__Check("1\n2\n3");
