// vybe-test: kotlin/kotlin_arrays_creation/test_double_array_range_with_fill
// origin: languages/kotlin/tests/kotlin/test_kotlin_arrays_creation.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = DoubleArray(3)
            values.fill(2.5)
            __check((values.sum()).toString(), "7.5")
        }
