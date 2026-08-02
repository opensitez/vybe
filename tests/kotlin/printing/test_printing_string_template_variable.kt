// vybe-test: kotlin/printing/test_printing_string_template_variable
// origin: languages/kotlin/tests/kotlin/test_printing.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val name = "kotlin"
            __check(("language=$name").toString(), "language=kotlin")
        }
