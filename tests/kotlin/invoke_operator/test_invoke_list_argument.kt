// vybe-test: kotlin/invoke_operator/test_invoke_list_argument
// origin: languages/kotlin/tests/kotlin/test_invoke_operator.rs

class Joiner {
            operator fun invoke(parts: List<String>): String = parts.joinToString(":")
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Joiner()(listOf("a", "b"))).toString(), "a:b")
        }
