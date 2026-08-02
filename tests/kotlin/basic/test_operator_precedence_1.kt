// vybe-test: kotlin/basic/test_operator_precedence_1
// origin: languages/kotlin/tests/kotlin/test_basic.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val res = 2 + 3 * 4
            __check((res).toString(), "14")
        }
