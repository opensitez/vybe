// vybe-test: kotlin/kotlin_set_apis/test_set_of_not_in
// origin: languages/kotlin/tests/kotlin/test_kotlin_set_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val set = setOf(1, 2, 3)
            __check((4 !in set).toString(), "true")
            __check((2 !in set).toString(), "false")
        }
