// vybe-test: csharp/csharp_generics/generic_list_usage
// origin: languages/csharp/tests/csharp/test_csharp_generics.rs

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
list.Add(10);
list.Add(20);
list.Add(30);
__P((list.Count).ToString());
__P((list[1]).ToString());
__Check("3\n20");
