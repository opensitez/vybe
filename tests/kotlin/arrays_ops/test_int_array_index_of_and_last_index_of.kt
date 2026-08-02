// vybe-test: kotlin/arrays_ops/test_int_array_index_of_and_last_index_of
// origin: languages/kotlin/tests/kotlin/test_arrays_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val nums = intArrayOf(4, 1, 4, 2, 4)
            __check((nums.indexOf(4)).toString(), "0")
            __check((nums.lastIndexOf(4)).toString(), "4")
            __check((nums.indexOf(9)).toString(), "-1")
        }
