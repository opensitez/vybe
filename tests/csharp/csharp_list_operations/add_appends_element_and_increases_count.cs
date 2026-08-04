// vybe-test: csharp/csharp_list_operations/add_appends_element_and_increases_count
// origin: languages/csharp/tests/csharp/test_csharp_list_operations.rs

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

var list = new System.Collections.Generic.List<int>();
list.Add(1); list.Add(2);
__P((list.Count).ToString());
__Check("2");
