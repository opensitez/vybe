// vybe-test: kotlin/mutable_set_apis/test_mutable_set_contains_all
// origin: languages/kotlin/tests/kotlin/test_mutable_set_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = mutableSetOf(1, 2, 3)
            __check((values.containsAll(listOf(1, 3))).toString(), "true")
            __check((values.containsAll(listOf(1, 4))).toString(), "false")
        }
