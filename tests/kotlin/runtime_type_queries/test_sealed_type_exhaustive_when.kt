// vybe-test: kotlin/runtime_type_queries/test_sealed_type_exhaustive_when
// origin: languages/kotlin/tests/kotlin/test_runtime_type_queries.rs

sealed interface Node
        data class Leaf(val value: Int) : Node
        data class Branch(val left: Node, val right: Node) : Node

        fun classify(node: Node): String = when (node) {
            is Leaf -> "leaf"
            is Branch -> "branch"
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = Leaf(1)
            val b = Branch(Leaf(2), Leaf(3))
            __check((classify(a)).toString(), "leaf")
            __check((classify(b)).toString(), "branch")
        }
