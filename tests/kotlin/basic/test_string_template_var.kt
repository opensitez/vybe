// vybe-test: kotlin/basic/test_string_template_var
// origin: languages/kotlin/tests/kotlin/test_basic.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val name = "Vybe"
            __check(("Hello $name").toString(), "Hello Vybe")
        }
