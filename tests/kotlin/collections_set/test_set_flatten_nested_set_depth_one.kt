// vybe-test: kotlin/collections_set/test_set_flatten_nested_set_depth_one
// origin: languages/kotlin/tests/kotlin/test_collections_set.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val nested: Set<Set<Int>> = setOf(setOf(1, 2), setOf(3, 4))
            val flattened = nested.flatten().toSet()
            __check((flattened.size).toString(), "4")
            __check((flattened.contains(3)).toString(), "true")
        }
