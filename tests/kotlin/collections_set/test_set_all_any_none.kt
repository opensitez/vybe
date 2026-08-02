// vybe-test: kotlin/collections_set/test_set_all_any_none
// origin: languages/kotlin/tests/kotlin/test_collections_set.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = setOf(2, 4, 6)
            __check((values.all { it % 2 == 0 }).toString(), "true")
            __check((values.any { it > 5 }).toString(), "true")
            __check((values.none { it < 0 }).toString(), "true")
            __check((values.any { it > 10 }).toString(), "false")
        }
