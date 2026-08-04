// vybe-test: csharp/csharp_list_operations/sort_orders_elements_ascending
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

var list = new System.Collections.Generic.List<int>{3,1,2};
list.Sort();
__P((list[0]).ToString()); __P((list[2]).ToString());
__Check("1\n3");
