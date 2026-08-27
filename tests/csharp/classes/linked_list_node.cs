// vybe-test: csharp/classes/linked_list_node
// origin: languages/csharp/tests/csharp/test_classes.rs

using static __Harness;

var a = new Node(1);
var b = new Node(2);
var c = new Node(3);
a.next = b;
b.next = c;
__P((a.value).ToString());
__P((a.next.value).ToString());
__P((a.next.next.value).ToString());
__Check("1\n2\n3");

class Node {
            public int value;
            public Node next;
            public Node(int v) { this.value = v; this.next = null; }
        }

public static class __Harness {
    public static string __buf = "";
    public static void __P(string s) { __buf = __buf + s + "\n"; }
    public static void __Pr(string s) { __buf = __buf + s; }
    public static void __Check(string want) {
        if (__buf != want && __buf != want + "\n") {
            Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
            throw new Exception("assertion failed");
        }
    }
}
