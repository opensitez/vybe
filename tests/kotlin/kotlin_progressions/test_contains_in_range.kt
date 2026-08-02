// vybe-test: kotlin/kotlin_progressions/test_contains_in_range
// origin: languages/kotlin/tests/kotlin/test_kotlin_progressions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check(((5 in 1..10).toString()).toString(), "true")
            __check(((11 in 1..10).toString()).toString(), "false")
            __check(((1L in 1L..10L).toString()).toString(), "true")
            __check(((10L in 1L until 10L).toString()).toString(), "false")
        }
