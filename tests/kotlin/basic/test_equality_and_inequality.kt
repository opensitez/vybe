// vybe-test: kotlin/basic/test_equality_and_inequality
// origin: languages/kotlin/tests/kotlin/test_basic.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check(("a" == "a").toString(), "true")
            __check(("a" != "b").toString(), "true")
            __check((10 <= 10).toString(), "true")
            __check((9 >= 10).toString(), "false")
        }
