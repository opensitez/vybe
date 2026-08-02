// vybe-test: kotlin/mutable_list_apis/test_mutable_list_contains_checks
// origin: languages/kotlin/tests/kotlin/test_mutable_list_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = mutableListOf("a", "b", "c")
            __check((values.contains("b")).toString(), "true")
            __check((values.isNotEmpty()).toString(), "true")
            __check((values.none()).toString(), "false")
        }
