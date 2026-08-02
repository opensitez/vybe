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

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val root = Node.Branch(Node.Leaf(1), Node.Leaf(2))
            __check((count(root)).toString(), "3")
        }
