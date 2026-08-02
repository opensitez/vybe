// vybe-test: kotlin/if_expressions/test_if_on_collections
// origin: languages/kotlin/tests/kotlin/test_if_expressions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = listOf(1, 2)
            val size = if (values.isEmpty()) 0 else values.size
            val first = if (values.isNotEmpty()) values[0] else -1
            __check((size).toString(), "2")
            __check((first).toString(), "1")
        }
