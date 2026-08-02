// vybe-test: kotlin/collections/test_array_of_nullable_length_three
// origin: languages/kotlin/tests/kotlin/test_collections.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values: Array<Int?> = Array(3) { null }
            values[2] = 14
            __check((values[0] == null).toString(), "true")
            __check((values[2] + 1).toString(), "15")
        }
