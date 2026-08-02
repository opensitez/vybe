// vybe-test: kotlin/collections_set/test_empty_set_intersection
// origin: languages/kotlin/tests/kotlin/test_collections_set.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val left = emptySet<Int>()
            val right = setOf(1, 2)
            val result = left.intersect(right)
            __check((result.isEmpty()).toString(), "true")
            __check((result.size).toString(), "0")
        }
