// vybe-test: kotlin/arrays_ops/test_int_array_get_or_else_get_or_null
// origin: languages/kotlin/tests/kotlin/test_arrays_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val nums = intArrayOf(9, 8, 7)
            __check((nums.getOrElse(1) { -1 }).toString(), "8")
            __check((nums.getOrNull(10) ?: -1).toString(), "-1")
        }
