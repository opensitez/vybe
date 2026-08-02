// vybe-test: kotlin/scope_shadowing/test_receiver_shadowing_in_with
// origin: languages/kotlin/tests/kotlin/test_scope_shadowing.rs

data class Node(val label: String)
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val node = Node("outer")
            val label = "local"
            val out = with(node) {
                val label = "with"
                label + "|" + this.label
            }
            __check((out).toString(), "with|outer")
            __check((label).toString(), "local")
        }
