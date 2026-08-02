// vybe-test: kotlin/kotlin_string_find_contains/test_contains_none
// origin: languages/kotlin/tests/kotlin/test_kotlin_string_find_contains.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val s = ""
            __check((s.contains("a").toString()).toString(), "false")
            __check((s.isEmpty().toString()).toString(), "true")
        }
