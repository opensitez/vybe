// vybe-test: kotlin/arrays_ops/test_array_zip_with_prefix_and_length_mismatch
// origin: languages/kotlin/tests/kotlin/test_arrays_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val left = intArrayOf(1, 2, 3)
            val right = intArrayOf(10, 20)
            val pairs = left.zip(right.toTypedArray())
            val values = pairs.joinToString("|") { it.first.toString() + ":" + it.second.toString() }
            __check((values).toString(), "1:10|2:20")
        }
