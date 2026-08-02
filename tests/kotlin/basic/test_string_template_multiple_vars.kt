// vybe-test: kotlin/basic/test_string_template_multiple_vars
// origin: languages/kotlin/tests/kotlin/test_basic.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val lang = "Kotlin"
            val ver = "1.9"
            __check(("$lang $ver").toString(), "Kotlin 1.9")
        }
