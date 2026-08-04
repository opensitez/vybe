// vybe-test: csharp/csharp_bcl_collections/linked_list_add_first_inserts_before_current_head
// origin: languages/csharp/tests/csharp/test_csharp_bcl_collections.rs

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

var list = new System.Collections.Generic.LinkedList<int>();
list.AddLast(2);
list.AddFirst(1);
__P((list.First.Value).ToString());
__Check("1");
