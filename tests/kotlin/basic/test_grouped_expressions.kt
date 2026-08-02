// vybe-test: kotlin/basic/test_grouped_expressions
// origin: languages/kotlin/tests/kotlin/test_basic.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = (10 + 20) * (30 - 10)
            __check((a).toString(), "600")
        }
