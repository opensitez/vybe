// vybe-test: kotlin/recursion/test_recursion_tree_depth
// origin: languages/kotlin/tests/kotlin/test_recursion.rs

class Node(val left: Node?, val right: Node?, val value: Int)
        fun depth(node: Node?): Int = if (node == null) 0 else 1 + maxOf(depth(node.left), depth(node.right))
        fun maxOf(a: Int, b: Int): Int = if (a > b) a else b
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val t = Node(Node(null, null, 2), null, 1)
            __check((depth(t)).toString(), "2")
        }
