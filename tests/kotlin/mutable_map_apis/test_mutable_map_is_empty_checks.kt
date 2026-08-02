// vybe-test: kotlin/mutable_map_apis/test_mutable_map_is_empty_checks
// origin: languages/kotlin/tests/kotlin/test_mutable_map_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = mutableMapOf<String, Int>()
            __check((values.isEmpty()).toString(), "true")
            values["a"] = 1
            __check((values.isNotEmpty()).toString(), "true")
        }
