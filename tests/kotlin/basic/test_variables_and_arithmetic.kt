// vybe-test: kotlin/basic/test_variables_and_arithmetic
// origin: languages/kotlin/tests/kotlin/test_basic.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = 15
            var b = 25
            val c = a + b
            __check((c).toString(), "40")
        }
