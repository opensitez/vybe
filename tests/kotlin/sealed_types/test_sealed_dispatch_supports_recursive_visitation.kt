// vybe-test: kotlin/sealed_types/test_sealed_dispatch_supports_recursive_visitation
// origin: languages/kotlin/tests/kotlin/test_sealed_types.rs

sealed class Node {
            class Leaf(val value: Int) : Node()
            class Branch(val left: Node, val right: Node) : Node()
        }

        fun sum(node: Node): Int {
            return when (node) {
                is Node.Leaf -> node.value
                is Node.Branch -> sum(node.left) + sum(node.right)
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val tree = Node.Branch(Node.Leaf(1), Node.Branch(Node.Leaf(2), Node.Leaf(3)))
            __check((sum(tree)).toString(), "6")
        }
