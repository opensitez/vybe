// vybe-test: kotlin/kotlin_string_find_contains/test_match_indexes
// origin: languages/kotlin/tests/kotlin/test_kotlin_string_find_contains.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val s = "x-ay-x"
            __check((s.indexOf("-").toString()).toString(), "1")
            __check((s.lastIndexOf("x").toString()).toString(), "4")
        }
