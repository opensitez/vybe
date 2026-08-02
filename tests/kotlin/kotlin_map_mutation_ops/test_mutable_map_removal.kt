// vybe-test: kotlin/kotlin_map_mutation_ops/test_mutable_map_removal
// origin: languages/kotlin/tests/kotlin/test_kotlin_map_mutation_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val m = mutableMapOf("x" to 1, "y" to 2)
            m.remove("x")
            __check((m.containsKey("x")).toString(), "false")
            __check((m.size).toString(), "1")
        }
