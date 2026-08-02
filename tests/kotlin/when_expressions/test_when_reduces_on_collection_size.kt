// vybe-test: kotlin/when_expressions/test_when_reduces_on_collection_size
// origin: languages/kotlin/tests/kotlin/test_when_expressions.rs

fun sizeLabel(values: List<Int>): String {
            return when (values.size) {
                0 -> "empty"
                in 1..2 -> "small"
                in 3..4 -> "mid"
                else -> "large"
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((sizeLabel(listOf())).toString(), "empty")
            __check((sizeLabel(listOf(1))).toString(), "small")
            __check((sizeLabel(listOf(1, 2, 3))).toString(), "mid")
            __check((sizeLabel(listOf(1, 2, 3, 4, 5))).toString(), "large")
        }
