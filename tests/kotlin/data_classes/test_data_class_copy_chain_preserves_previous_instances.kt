// vybe-test: kotlin/data_classes/test_data_class_copy_chain_preserves_previous_instances
// origin: languages/kotlin/tests/kotlin/test_data_classes.rs

data class Node(val id: Int, val label: String)

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = Node(1, "a")
            val b = a.copy(label = "b")
            val c = b.copy(id = 3)
            __check((a.label).toString(), "a")
            __check((b.id).toString(), "1")
            __check((c.label).toString(), "b")
            __check((a == b).toString(), "false")
            __check((b == c).toString(), "false")
        }
