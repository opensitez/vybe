// vybe-test: kotlin/kotlin_map_query_ops/test_map_merging_plus
// origin: languages/kotlin/tests/kotlin/test_kotlin_map_query_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = mapOf("a" to 1, "b" to 2)
            val b = mapOf("b" to 3, "c" to 4)
            val merged = a + b
            __check((merged["a"].toString()).toString(), "1")
            __check((merged["b"].toString()).toString(), "3")
            __check((merged["c"].toString()).toString(), "4")
        }
