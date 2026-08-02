// vybe-test: kotlin/mutable_map_apis/test_mutable_map_remove_key_value_mismatch
// origin: languages/kotlin/tests/kotlin/test_mutable_map_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = mutableMapOf("a" to 1)
            val removed = values.remove("a", 9)
            __check((removed).toString(), "false")
            __check((values.size).toString(), "1")
            __check((values["a"]).toString(), "1")
        }
