// vybe-test: kotlin/arrays_ops/test_array_with_function_mapping
// origin: languages/kotlin/tests/kotlin/test_arrays_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val nums = intArrayOf(1, 2, 3).map { it * 2 }.toIntArray()
            __check((nums.joinToString(",")).toString(), "2,4,6")
        }
