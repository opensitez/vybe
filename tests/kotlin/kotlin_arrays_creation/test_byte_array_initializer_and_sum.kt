// vybe-test: kotlin/kotlin_arrays_creation/test_byte_array_initializer_and_sum
// origin: languages/kotlin/tests/kotlin/test_kotlin_arrays_creation.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val bytes = byteArrayOf(1, 2, 3)
            __check((bytes[1] + bytes[2]).toString(), "5")
        }
