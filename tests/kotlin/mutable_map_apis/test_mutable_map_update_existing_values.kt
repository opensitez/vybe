// vybe-test: kotlin/mutable_map_apis/test_mutable_map_update_existing_values
// origin: languages/kotlin/tests/kotlin/test_mutable_map_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = mutableMapOf("a" to 1, "b" to 2)
            values["a"] = 4
            values.put("b", 5)
            __check((values["a"]).toString(), "4")
            __check((values["b"]).toString(), "5")
        }
