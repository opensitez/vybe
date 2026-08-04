// vybe-test: csharp/csharp_list_operations/remove_deletes_first_matching_element
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

var list = new System.Collections.Generic.List<int>{1,2,2,3};
list.Remove(2);
__P((list.Count).ToString()); __P((list[1]).ToString());
__Check("3\n2");
