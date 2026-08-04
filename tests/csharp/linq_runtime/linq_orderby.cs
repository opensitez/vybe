// vybe-test: csharp/linq_runtime/linq_orderby
// origin: languages/csharp/tests/csharp/test_linq_runtime.rs

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

var list = new List<int>();
list.Add(5); list.Add(3); list.Add(1); list.Add(4); list.Add(2);
list.OrderBy(x => x).ForEach(x => __P((x).ToString()));
__Check("1\n2\n3\n4\n5");
