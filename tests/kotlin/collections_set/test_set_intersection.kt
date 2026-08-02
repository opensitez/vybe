// vybe-test: kotlin/collections_set/test_set_intersection
// origin: languages/kotlin/tests/kotlin/test_collections_set.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = setOf(1, 2, 3)
            val b = setOf(2, 4, 3)
            val overlap = a intersect b
            __check((overlap.size).toString(), "2")
            __check((overlap.contains(2)).toString(), "true")
            __check((overlap.contains(4)).toString(), "false")
        }
