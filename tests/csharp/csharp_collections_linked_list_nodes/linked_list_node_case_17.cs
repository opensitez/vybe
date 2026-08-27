// vybe-test: csharp/csharp_collections_linked_list_nodes/linked_list_node_case_17

string __buf = "";

void __P(string s) {
    __buf = __buf + s + "\n";
}

void __Pr(string s) {
    __buf = __buf + s;
}

void __Check(string want) {
    if (__buf != want && __buf != want + "\n") {
        Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
        throw new Exception("assertion failed");
    }
}

var list = new System.Collections.Generic.LinkedList<string>();
var mid = list.AddLast("Mid_17");
list.AddBefore(mid, "First_17");
list.AddAfter(mid, "Last_17");
__P(list.First.Value);
__P(list.Last.Value);
__Check("First_17\nLast_17");
