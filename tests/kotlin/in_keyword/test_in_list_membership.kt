// vybe-test: kotlin/in_keyword/test_in_list_membership
// origin: languages/kotlin/tests/kotlin/test_in_keyword.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = listOf("a", "b", "c")
            __check(("b" in values).toString(), "true")
            __check(("z" !in values).toString(), "true")
        }
