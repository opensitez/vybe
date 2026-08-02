// vybe-test: kotlin/mutable_map_apis/test_mutable_map_contains_key_and_value
// origin: languages/kotlin/tests/kotlin/test_mutable_map_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = mutableMapOf("a" to 1, "b" to 2)
            __check((values.containsKey("a")).toString(), "true")
            __check((values.containsValue(2)).toString(), "true")
            __check((values.containsValue(3)).toString(), "false")
        }
