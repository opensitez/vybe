// vybe-test: kotlin/arrays_ops/test_concatenate_arrays_via_plus
// origin: languages/kotlin/tests/kotlin/test_arrays_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = intArrayOf(1, 2)
            val b = intArrayOf(3, 4)
            val c = a + b
            __check((c.joinToString(",")).toString(), "1,2,3,4")
        }
