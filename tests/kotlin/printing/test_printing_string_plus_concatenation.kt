// vybe-test: kotlin/printing/test_printing_string_plus_concatenation
// origin: languages/kotlin/tests/kotlin/test_printing.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val prefix = "a"
            __check((prefix + " + " + 2).toString(), "a + 2")
        }
