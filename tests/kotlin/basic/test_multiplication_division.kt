// vybe-test: kotlin/basic/test_multiplication_division
// origin: languages/kotlin/tests/kotlin/test_basic.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = 12
            val b = 4
            __check((a * b).toString(), "48")
            __check((a / b).toString(), "3")
        }
