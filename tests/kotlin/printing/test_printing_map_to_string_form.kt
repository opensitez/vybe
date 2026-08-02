// vybe-test: kotlin/printing/test_printing_map_to_string_form
// origin: languages/kotlin/tests/kotlin/test_printing.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val data = linkedMapOf("a" to 1, "b" to 2)
            __check((data).toString(), "{a=1, b=2}")
        }
