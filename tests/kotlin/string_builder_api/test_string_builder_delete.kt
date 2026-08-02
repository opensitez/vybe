// vybe-test: kotlin/string_builder_api/test_string_builder_delete
// origin: languages/kotlin/tests/kotlin/test_string_builder_api.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val out = StringBuilder("abcd")
            out.delete(1, 3)
            __check((out.toString()).toString(), "ad")
        }
