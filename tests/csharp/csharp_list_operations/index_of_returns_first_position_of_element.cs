// vybe-test: csharp/csharp_list_operations/index_of_returns_first_position_of_element
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

var list = new System.Collections.Generic.List<int>{5,10,5};
__P((list.IndexOf(5)).ToString());
__Check("0");
