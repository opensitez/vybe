// vybe-test: kotlin/arrays_ops/test_array_any_all_none_count
// origin: languages/kotlin/tests/kotlin/test_arrays_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val nums = intArrayOf(1, 2, 3, 4)
            __check((nums.any { it > 3 }).toString(), "true")
            __check((nums.all { it > 0 }).toString(), "true")
            __check((nums.none { it < 0 }).toString(), "true")
            __check((nums.count { it % 2 == 0 }).toString(), "2")
        }
