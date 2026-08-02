// vybe-test: kotlin/block_expressions/test_block_inside_when_subject
// origin: languages/kotlin/tests/kotlin/test_block_expressions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
        val x = when (run { 3 }) {
            run { 1 + 2 } -> "a"
            else -> "b"
        }
        __check((x).toString(), "a")
    }
