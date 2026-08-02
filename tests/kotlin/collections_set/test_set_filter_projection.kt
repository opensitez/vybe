// vybe-test: kotlin/collections_set/test_set_filter_projection
// origin: languages/kotlin/tests/kotlin/test_collections_set.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = setOf(1, 2, 3, 4, 5)
            val evens = values.filter { it % 2 == 0 }
            __check((evens.size).toString(), "2")
            __check((evens[1]).toString(), "4")
        }
