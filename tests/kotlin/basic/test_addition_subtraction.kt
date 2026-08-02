// vybe-test: kotlin/basic/test_addition_subtraction
// origin: languages/kotlin/tests/kotlin/test_basic.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = 50
            val b = 20
            __check((a + b).toString(), "70")
            __check((a - b).toString(), "30")
        }
