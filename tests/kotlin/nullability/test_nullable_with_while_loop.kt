// vybe-test: kotlin/nullability/test_nullable_with_while_loop
// origin: languages/kotlin/tests/kotlin/test_nullability.rs

class Node {
            var next: Node? = null
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val head: Node? = Node()
            val result = if (head?.next == null) "end" else "mid"
            __check((result).toString(), "end")
        }
