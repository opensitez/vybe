// vybe-test: kotlin/mutable_set_apis/test_mutable_set_contains_none
// origin: languages/kotlin/tests/kotlin/test_mutable_set_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = mutableSetOf("a", "b")
            __check((values.contains("c")).toString(), "false")
            __check((values.containsAll(listOf("a", "b"))).toString(), "true")
        }
