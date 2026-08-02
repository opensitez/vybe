// vybe-test: kotlin/data_classes/test_data_class_copy_chain_preserves_immutability_surface
// origin: languages/kotlin/tests/kotlin/test_data_classes.rs

data class Node(val id: Int, val next: Node?)

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = Node(1, Node(2, null))
            val b = a.copy(id = 10)
            __check((a.id).toString(), "1")
            __check((a.next?.id).toString(), "2")
            __check((b.id).toString(), "10")
            __check((b.next?.id).toString(), "2")
        }
