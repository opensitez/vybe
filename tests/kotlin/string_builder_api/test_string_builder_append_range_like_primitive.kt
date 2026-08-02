// vybe-test: kotlin/string_builder_api/test_string_builder_append_range_like_primitive
// origin: languages/kotlin/tests/kotlin/test_string_builder_api.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val out = StringBuilder()
            val chars = charArrayOf('x', 'y', 'z')
            out.append(chars)
            __check((out.toString()).toString(), "xyz")
        }
