// vybe-test: kotlin/kotlin_iterable_to_collections/test_map_not_null_to_map
// origin: languages/kotlin/tests/kotlin/test_kotlin_iterable_to_collections.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = listOf("a" to 1, "b" to 2)
            val keys = values.map { it.first }
            val valuesOut = values.map { it.second }
            __check((keys.joinToString(",")).toString(), "a,b")
            __check((valuesOut.joinToString(",")).toString(), "1,2")
        }
