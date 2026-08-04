// vybe-test: kotlin/sealed_types/test_sealed_class_with_leaf_subclass_instances
// origin: languages/kotlin/tests/kotlin/test_sealed_types.rs

sealed class Node {
            class Leaf(val value: Int) : Node()
            class Branch(val left: Node, val right: Node) : Node()
        }

        fun count(node: Node): Int {
            return when (node) {
                is Node.Leaf -> 1
                is Node.Branch -> 1 + count(node.left) + count(node.right)
            }
        }

        var __buf: String = ""

fun __p(s: String) {
    __buf = __buf + s + "\n"
}

fun __pr(s: String) {
    __buf = __buf + s
}

// The final `println` contributes a trailing newline that the expected line
// vector never carried, so BOTH forms are accepted. Written as two equality
// tests rather than trimming: `String.endsWith` is not implemented in Vybe's
// Kotlin (measured — `"ab\n".endsWith("\n")` throws "undefined is not
// callable"), and a harness that cannot run asserts nothing at all. The cargo
// helper split on "\n" and popped trailing empties, so the two forms were
// equivalent there too.
fun __check(want: String) {
    if (__buf != want && __buf != want + "\n") {
        println("FAIL: want [" + want + "] got [" + __buf + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val root = Node.Branch(Node.Leaf(1), Node.Leaf(2))
            __p((count(root)).toString())
        
__check("3")
}
