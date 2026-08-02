// vybe-test: kotlin/kotlin_arrays_creation/test_copy_of_range_with_offseted_start
// origin: languages/kotlin/tests/kotlin/test_kotlin_arrays_creation.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = intArrayOf(0, 1, 2, 3, 4, 5)
            val slice = values.copyOfRange(1, 4)
            __check((slice.joinToString(",")).toString(), "1,2,3")
        }
