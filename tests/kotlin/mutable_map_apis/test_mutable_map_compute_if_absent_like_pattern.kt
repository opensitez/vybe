// vybe-test: kotlin/mutable_map_apis/test_mutable_map_compute_if_absent_like_pattern
// origin: languages/kotlin/tests/kotlin/test_mutable_map_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = mutableMapOf("a" to 1)
            if (!values.containsKey("b")) {
                values["b"] = values.size
            }
            __check((values["b"]).toString(), "1")
            __check((values["a"]).toString(), "1")
        }
