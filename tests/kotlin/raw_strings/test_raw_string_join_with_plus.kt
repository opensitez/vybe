// vybe-test: kotlin/raw_strings/test_raw_string_join_with_plus
// origin: languages/kotlin/tests/kotlin/test_raw_strings.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = """a"""
            val b = """b"""
            __check((a + b).toString(), "ab")
        }
