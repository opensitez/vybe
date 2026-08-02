// vybe-test: kotlin/kotlin_map_mutation_ops/test_mutable_map_put_and_update
// origin: languages/kotlin/tests/kotlin/test_kotlin_map_mutation_ops.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val m = mutableMapOf("a" to 1)
            m["a"] = 2
            m.put("b", 3)
            __check((m["a"]).toString(), "2")
            __check((m["b"]).toString(), "3")
        }
