// vybe-test: kotlin/in_keyword/test_in_set_lookup
// origin: languages/kotlin/tests/kotlin/test_in_keyword.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = setOf(1, 2, 3)
            __check((2 in values).toString(), "true")
            __check((8 in values).toString(), "false")
        }
