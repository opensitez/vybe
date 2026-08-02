// vybe-test: kotlin/basic/test_modulo_operator
// origin: languages/kotlin/tests/kotlin/test_basic.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = 17
            val b = 5
            __check((a % b).toString(), "2")
        }
