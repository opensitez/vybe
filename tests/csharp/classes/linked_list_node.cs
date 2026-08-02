// vybe-test: csharp/classes/linked_list_node
// origin: languages/csharp/tests/csharp/test_classes.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
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
        __Check((a.value).ToString(), "1");
        __Check((a.next.value).ToString(), "2");
        __Check((a.next.next.value).ToString(), "3");
