// vybe-test: kotlin/kotlin_map_query_ops/test_map_map_keys_values
// origin: languages/kotlin/tests/kotlin/test_kotlin_map_query_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val m = mapOf("a" to 1, "b" to 2)
            val mappedKeys = m.mapKeys { it.key.uppercase() }
            val mappedValues = m.mapValues { it.value + 10 }
            __check((mappedKeys["A"].toString()).toString(), "1")
            __check((mappedValues["b"].toString()).toString(), "12")
        }
