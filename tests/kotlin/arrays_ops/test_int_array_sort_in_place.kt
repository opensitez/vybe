// vybe-test: kotlin/arrays_ops/test_int_array_sort_in_place
// origin: languages/kotlin/tests/kotlin/test_arrays_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val nums = intArrayOf(3, 1, 4, 1, 5, 9)
            nums.sort()
            __check((nums.joinToString(",")).toString(), "1,1,3,4,5,9")
        }
