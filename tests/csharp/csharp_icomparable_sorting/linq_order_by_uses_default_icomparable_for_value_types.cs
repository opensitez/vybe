// vybe-test: csharp/csharp_icomparable_sorting/linq_order_by_uses_default_icomparable_for_value_types
// origin: languages/csharp/tests/csharp/test_csharp_icomparable_sorting.rs

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

var result = new[]{3,1,2}.OrderBy(x=>x);
foreach(var n in result) __P((n).ToString());
__Check("1\n2\n3");
