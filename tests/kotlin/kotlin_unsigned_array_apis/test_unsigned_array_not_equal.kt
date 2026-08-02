// vybe-test: kotlin/kotlin_unsigned_array_apis/test_unsigned_array_not_equal
// origin: languages/kotlin/tests/kotlin/test_kotlin_unsigned_array_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = ulongArrayOf(1uL, 2uL)
            val b = ulongArrayOf(1uL, 3uL)
            __check(((a == b).toString()).toString(), "false")
        }
