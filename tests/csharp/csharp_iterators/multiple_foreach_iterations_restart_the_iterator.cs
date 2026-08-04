// vybe-test: csharp/csharp_iterators/multiple_foreach_iterations_restart_the_iterator
// origin: languages/csharp/tests/csharp/test_csharp_iterators.rs

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

System.Collections.Generic.IEnumerable<int> Three() {
    yield return 1; yield return 2; yield return 3;
}
int total=0;
foreach(var x in Three()) total+=x;
foreach(var x in Three()) total+=x;
__P((total).ToString());
__Check("12");
