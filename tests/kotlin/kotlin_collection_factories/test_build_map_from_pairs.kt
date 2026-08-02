// vybe-test: kotlin/kotlin_collection_factories/test_build_map_from_pairs
// origin: languages/kotlin/tests/kotlin/test_kotlin_collection_factories.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = buildMap {
                put(1, "a")
                put(2, "b")
                put(1, "c")
            }
            __check((value.size).toString(), "2")
            __check((value[1]).toString(), "c")
            __check((value[2]).toString(), "b")
        }
