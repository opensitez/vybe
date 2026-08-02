// vybe-test: kotlin/range_projection/test_when_with_range_subject
// origin: languages/kotlin/tests/kotlin/test_range_projection.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val v = 4
            val label = when (v) {
                in 1..3 -> "small"
                in 4..6 -> "mid"
                else -> "big"
            }
            __check((label).toString(), "mid")
            __check((v in 1..10).toString(), "true")
        }
