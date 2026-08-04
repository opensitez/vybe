// vybe-test: csharp/classes/linked_list_node
// origin: languages/csharp/tests/csharp/test_classes.rs

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

class Node {
            public int value;
            public Node next;
            public Node(int v) { this.value = v; this.next = null; }
        }
        var a = new Node(1);
        var b = new Node(2);
        var c = new Node(3);
        a.next = b;
        b.next = c;
        __P((a.value).ToString());
        __P((a.next.value).ToString());
        __P((a.next.next.value).ToString());
__Check("1\n2\n3");
