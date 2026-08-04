// vybe-test: csharp/linq_runtime/linq_chain_where_foreach
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

var nums = new List<int>();
nums.Add(1); nums.Add(2); nums.Add(3); nums.Add(4); nums.Add(5);
nums.Where(x => x > 3).ForEach(x => __P((x).ToString()));
__Check("4\n5");
