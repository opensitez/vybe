// vybe-test: kotlin/member_references/test_unbound_property_reference_in_sorting
// origin: languages/kotlin/tests/kotlin/test_member_references.rs

data class Node(val score: Int)
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val nodes = listOf(Node(2), Node(1), Node(3))
            val out = nodes.sortedBy(Node::score).joinToString(",") { it.score.toString() }
            __check((out).toString(), "1,2,3")
        }
