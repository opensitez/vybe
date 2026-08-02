// vybe-test: kotlin/primitive_array_apis/test_long_array_min_max
// origin: languages/kotlin/tests/kotlin/test_primitive_array_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = longArrayOf(10L, 3L, 7L)
            __check((values.minOrNull()).toString(), "3")
            __check((values.maxOrNull()).toString(), "10")
        }
