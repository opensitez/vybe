// vybe-test: kotlin/kotlin_arrays_creation/test_array_constructor_lambda_and_copy
// origin: languages/kotlin/tests/kotlin/test_kotlin_arrays_creation.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = Array(4) { it * 2 }
            val copy = values.copyOf(2)
            __check((copy.joinToString(",")).toString(), "0,2")
        }
