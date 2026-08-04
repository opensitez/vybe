// vybe-test: csharp/csharp_collection_types/linked_list_add_after_inserts_between_nodes
// origin: languages/csharp/tests/csharp/test_csharp_collection_types.rs

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

var ll=new System.Collections.Generic.LinkedList<int>();
var n1=ll.AddFirst(1);
ll.AddAfter(n1,3);
ll.AddAfter(n1,2);
__P((ll.First.Next.Value).ToString());
__Check("2");
