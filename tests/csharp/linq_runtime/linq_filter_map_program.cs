// vybe-test: csharp/linq_runtime/linq_filter_map_program
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

var numbers = new List<int>();
numbers.Add(1); numbers.Add(2); numbers.Add(3); numbers.Add(4);
numbers.Add(5); numbers.Add(6); numbers.Add(7); numbers.Add(8);
var result = numbers.Where(n => n % 2 == 0).Select(n => n * n);
result.ForEach(x => __P((x).ToString()));
__Check("4\n16\n36\n64");
