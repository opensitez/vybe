// vybe-test: kotlin/kotlin_iterable_to_collections/test_to_map_from_pairs
// origin: languages/kotlin/tests/kotlin/test_kotlin_iterable_to_collections.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val pairs = listOf("a" to 1, "b" to 2)
            val out = pairs.toMap()
            __check((out.size).toString(), "2")
            __check((out["a"].toString()).toString(), "1")
            __check((out["b"].toString()).toString(), "2")
        }
