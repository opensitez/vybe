// vybe-test: kotlin/raw_strings/test_raw_string_concat_multiple
// origin: languages/kotlin/tests/kotlin/test_raw_strings.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val text = """a""" + """b""" + """c"""
            __check((text).toString(), "abc")
        }
