// vybe-test: kotlin/kotlin_multiline_strings/test_raw_string_join_with_pipe
// origin: languages/kotlin/tests/kotlin/test_kotlin_multiline_strings.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val lines = """a|b|c"""
            val parts = lines.split("|")
            __check((parts.size).toString(), "3")
            __check((parts.joinToString(",")).toString(), "a,b,c")
        }
