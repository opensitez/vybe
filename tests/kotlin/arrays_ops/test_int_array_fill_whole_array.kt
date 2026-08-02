// vybe-test: kotlin/arrays_ops/test_int_array_fill_whole_array
// origin: languages/kotlin/tests/kotlin/test_arrays_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val nums = IntArray(4)
            nums.fill(7)
            __check((nums.joinToString(",")).toString(), "7,7,7,7")
        }
