// vybe-test: kotlin/mutable_map_apis/test_mutable_map_put_all_values
// origin: languages/kotlin/tests/kotlin/test_mutable_map_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = mutableMapOf("a" to 1)
            values.putAll(mapOf("b" to 2, "c" to 3))
            __check((values.keys.joinToString(",")).toString(), "a,b,c")
            __check((values["b"]).toString(), "2")
            __check((values["c"]).toString(), "3")
        }
