// vybe-test: kotlin/range_projection/test_empty_range_from_invalid_step
// origin: languages/kotlin/tests/kotlin/test_range_projection.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val r = 5 downTo 10
            __check((r.isEmpty()).toString(), "true")
            __check((r.toList().isEmpty()).toString(), "true")
        }
