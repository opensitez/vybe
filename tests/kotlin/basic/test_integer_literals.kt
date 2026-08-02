// vybe-test: kotlin/basic/test_integer_literals
// origin: languages/kotlin/tests/kotlin/test_basic.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = 0
            val b = 42
            __check((a).toString(), "0")
            __check((b).toString(), "42")
        }
