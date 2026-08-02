// vybe-test: kotlin/sealed_types/test_sealed_nested_structure_can_be_traversed_iteratively
// origin: languages/kotlin/tests/kotlin/test_sealed_types.rs

sealed class Node {
            class Leaf(val value: Int) : Node()
            class Branch(val left: Node, val right: Node) : Node()
        }

        fun collect(node: Node): Int {
            return when (node) {
                is Node.Leaf -> 1
                is Node.Branch -> 1 + collect(node.left) + collect(node.right)
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val root = Node.Branch(
                Node.Leaf(1),
                Node.Branch(Node.Leaf(2), Node.Leaf(3))
            )
            __check((collect(root)).toString(), "4")
        }
