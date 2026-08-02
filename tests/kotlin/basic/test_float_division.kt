// vybe-test: kotlin/basic/test_float_division
// origin: languages/kotlin/tests/kotlin/test_basic.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val result = 7 / 2.0
            __check((result).toString(), "3.5")
        }
