// vybe-test: kotlin/string_builder_api/test_string_builder_clear_via_set_length_zero
// origin: languages/kotlin/tests/kotlin/test_string_builder_api.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val out = StringBuilder("clear")
            out.setLength(0)
            __check((out.isEmpty()).toString(), "true")
            __check((out.length).toString(), "0")
        }
