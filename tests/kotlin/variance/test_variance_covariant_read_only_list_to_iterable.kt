// vybe-test: kotlin/variance/test_variance_covariant_read_only_list_to_iterable
// origin: languages/kotlin/tests/kotlin/test_variance.rs

val values: List<String> = listOf("a", "b")
        val anyValues: List<Any> = values
        __check((anyValues.size).toString(), "2")

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}
