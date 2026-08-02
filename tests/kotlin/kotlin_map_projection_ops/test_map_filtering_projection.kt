// vybe-test: kotlin/kotlin_map_projection_ops/test_map_filtering_projection
// origin: languages/kotlin/tests/kotlin/test_kotlin_map_projection_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val m = mapOf("a" to 1, "b" to 3, "c" to 5)
            val f = m.filterValues { it > 1 }
            __check((f["b"]).toString(), "3")
            __check((f["a"]).toString(), "null")
        }
