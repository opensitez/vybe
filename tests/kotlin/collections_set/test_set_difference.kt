// vybe-test: kotlin/collections_set/test_set_difference
// origin: languages/kotlin/tests/kotlin/test_collections_set.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = setOf(1, 2, 3)
            val b = setOf(2, 4)
            val remaining = a - b
            __check((remaining.size).toString(), "2")
            __check((remaining.contains(2)).toString(), "false")
            __check((remaining.contains(1)).toString(), "true")
        }
