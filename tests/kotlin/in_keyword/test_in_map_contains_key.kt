// vybe-test: kotlin/in_keyword/test_in_map_contains_key
// origin: languages/kotlin/tests/kotlin/test_in_keyword.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val map = mapOf("a" to 1, "b" to 2)
            __check(("a" in map).toString(), "true")
            __check(("c" in map).toString(), "false")
            __check(("a" !in map).toString(), "false")
        }
