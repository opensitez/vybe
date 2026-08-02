// vybe-test: kotlin/kotlin_arrays_creation/test_int_array_of_constructor_and_indexing
// origin: languages/kotlin/tests/kotlin/test_kotlin_arrays_creation.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = intArrayOf(1, 2, 3, 4)
            __check((values[0] + values[3]).toString(), "5")
        }
