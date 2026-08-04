// vybe-test: csharp/csharp_map_set_collections/linked_list_add_first_and_add_last_preserve_order
// origin: languages/csharp/tests/csharp/test_csharp_map_set_collections.rs

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

using System.Collections.Generic; var items = new LinkedList<string>(); items.AddFirst("middle"); items.AddFirst("start"); items.AddLast("end"); foreach (var item in items) __P((item).ToString());
__Check("start\nmiddle\nend");
