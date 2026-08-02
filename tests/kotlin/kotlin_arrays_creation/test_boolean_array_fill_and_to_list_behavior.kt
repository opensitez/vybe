// vybe-test: kotlin/kotlin_arrays_creation/test_boolean_array_fill_and_to_list_behavior
// origin: languages/kotlin/tests/kotlin/test_kotlin_arrays_creation.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val bits = BooleanArray(4)
            bits[0] = true
            bits[2] = true
            __check((bits.count { it }).toString(), "2")
        }
