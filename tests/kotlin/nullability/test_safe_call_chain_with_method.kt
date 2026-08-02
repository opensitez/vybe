// vybe-test: kotlin/nullability/test_safe_call_chain_with_method
// origin: languages/kotlin/tests/kotlin/test_nullability.rs

class Node {
            fun name(): String {
                return "node"
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value: Node? = Node()
            __check((value?.name() ?: "none").toString(), "node")
        }
