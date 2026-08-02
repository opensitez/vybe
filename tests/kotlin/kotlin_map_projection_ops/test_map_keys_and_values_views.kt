// vybe-test: kotlin/kotlin_map_projection_ops/test_map_keys_and_values_views
// origin: languages/kotlin/tests/kotlin/test_kotlin_map_projection_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val m = mapOf("x" to 1, "y" to 2)
            __check((m.keys.toString()).toString(), "[x, y]")
            __check((m.values.toString()).toString(), "[1, 2]")
        }
