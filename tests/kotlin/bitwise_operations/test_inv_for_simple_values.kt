// vybe-test: kotlin/bitwise_operations/test_inv_for_simple_values
// origin: languages/kotlin/tests/kotlin/test_bitwise_operations.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((0.inv()).toString(), "-1")
            __check((1.inv()).toString(), "0")
            __check((255.inv()).toString(), "-256")
            __check((1023.inv()).toString(), "-1024")
        }
