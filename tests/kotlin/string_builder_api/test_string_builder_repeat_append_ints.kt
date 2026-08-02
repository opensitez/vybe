// vybe-test: kotlin/string_builder_api/test_string_builder_repeat_append_ints
// origin: languages/kotlin/tests/kotlin/test_string_builder_api.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val out = StringBuilder()
            repeat(3) { out.append('x') }
            __check((out.toString()).toString(), "xxx")
        }
