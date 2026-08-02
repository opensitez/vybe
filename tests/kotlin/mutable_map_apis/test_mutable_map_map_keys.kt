// vybe-test: kotlin/mutable_map_apis/test_mutable_map_map_keys
// origin: languages/kotlin/tests/kotlin/test_mutable_map_apis.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = mutableMapOf("a" to 1, "b" to 2)
            val mapped = values.mapKeys { it.key + it.value.toString() }
            __check((mapped.keys.joinToString(",")).toString(), "a1,b2")
        }
