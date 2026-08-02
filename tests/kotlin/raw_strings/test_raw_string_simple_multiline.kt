// vybe-test: kotlin/raw_strings/test_raw_string_simple_multiline
// origin: languages/kotlin/tests/kotlin/test_raw_strings.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val text = """line one
line two"""
            __check((text.lines().size).toString(), "2")
        }
