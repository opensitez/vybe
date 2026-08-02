// vybe-test: kotlin/preconditions/test_guarded_list_restriction
// origin: languages/kotlin/tests/kotlin/test_preconditions.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = listOf(1, 2, 3)
            require(values.isNotEmpty())
            check(values.size == 3)
            __check((values.sum()).toString(), "6")
        }
