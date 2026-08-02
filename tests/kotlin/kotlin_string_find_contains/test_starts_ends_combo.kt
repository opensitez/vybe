// vybe-test: kotlin/kotlin_string_find_contains/test_starts_ends_combo
// origin: languages/kotlin/tests/kotlin/test_kotlin_string_find_contains.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val s = "prefix_value_suffix"
            val starts = s.startsWith("prefix").toString()
            val ends = s.endsWith("suffix").toString()
            __check((starts).toString(), "true")
            __check((ends).toString(), "true")
        }
