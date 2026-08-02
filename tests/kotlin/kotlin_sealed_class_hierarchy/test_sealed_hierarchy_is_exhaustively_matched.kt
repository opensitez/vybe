// vybe-test: kotlin/kotlin_sealed_class_hierarchy/test_sealed_hierarchy_is_exhaustively_matched
// origin: languages/kotlin/tests/kotlin/test_kotlin_sealed_class_hierarchy.rs

sealed class Node {
            data class Value(val n: Int) : Node()
            data class Negate(val child: Node) : Node()
            data class Sum(val left: Node, val right: Node) : Node()
        }

        fun eval(node: Node): Int = when (node) {
            is Node.Value -> node.n
            is Node.Negate -> -eval(node.child)
            is Node.Sum -> eval(node.left) + eval(node.right)
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val expr = Node.Sum(Node.Value(3), Node.Negate(Node.Value(2)))
            __check((eval(expr)).toString(), "1")
        }
