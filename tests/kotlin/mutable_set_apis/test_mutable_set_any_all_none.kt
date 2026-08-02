// vybe-test: kotlin/mutable_set_apis/test_mutable_set_any_all_none
// origin: languages/kotlin/tests/kotlin/test_mutable_set_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = mutableSetOf(1, 2, 3)
            __check((values.any { it > 2 }).toString(), "true")
            __check((values.all { it > 0 }).toString(), "true")
            __check((values.none { it > 4 }).toString(), "true")
        }
