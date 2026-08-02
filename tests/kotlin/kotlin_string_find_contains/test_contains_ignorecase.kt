// vybe-test: kotlin/kotlin_string_find_contains/test_contains_ignorecase
// origin: languages/kotlin/tests/kotlin/test_kotlin_string_find_contains.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val s = "Case"
            __check((s.contains("c", ignoreCase = true).toString()).toString(), "true")
            __check((s.contains("X", ignoreCase = true).toString()).toString(), "false")
        }
