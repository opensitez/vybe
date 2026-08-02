// vybe-test: kotlin/type_cast_edges/test_cast_chain_on_map_values
// origin: languages/kotlin/tests/kotlin/test_type_cast_edges.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value: Any = mapOf("a" to "v")
            val map = value as? Map<String, String>
            __check((map?.get("a") ?: "none").toString(), "v")
        }
