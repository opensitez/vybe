// vybe-test: kotlin/type_casts/test_is_operator_true_branch
// origin: languages/kotlin/tests/kotlin/test_type_casts.rs

open class Node
        class Leaf : Node()

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val node: Node = Leaf()
            if (node is Node) {
                __check(("is_node").toString(), "is_node")
            }
        }
