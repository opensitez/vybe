// vybe-test: kotlin/string_builder_api/test_string_builder_append_format
// origin: languages/kotlin/tests/kotlin/test_string_builder_api.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val out = StringBuilder()
            val value = 7
            out.append("value=").append(value).append(",")
            __check((out.toString()).toString(), "value=7,")
        }
