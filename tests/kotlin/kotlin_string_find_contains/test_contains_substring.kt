// vybe-test: kotlin/kotlin_string_find_contains/test_contains_substring
// origin: languages/kotlin/tests/kotlin/test_kotlin_string_find_contains.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val s = "hello world"
            __check((s.contains("lo wo").toString()).toString(), "true")
            __check((s.contains("planet").toString()).toString(), "false")
        }
