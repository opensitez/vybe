// vybe-test: kotlin/arrays_ops/test_array_find_first_last_or_null
// origin: languages/kotlin/tests/kotlin/test_arrays_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val nums = intArrayOf(2, 4, 6)
            __check((nums.find { it > 3 } ?: -1).toString(), "4")
            __check((nums.findLast { it < 4 } ?: -1).toString(), "2")
            __check((nums.firstOrNull { it == 10 } ?: -1).toString(), "-1")
        }
