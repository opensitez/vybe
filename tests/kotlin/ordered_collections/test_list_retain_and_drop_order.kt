// vybe-test: kotlin/ordered_collections/test_list_retain_and_drop_order
// origin: languages/kotlin/tests/kotlin/test_ordered_collections.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = listOf(1, 2, 3, 4, 5)
            __check((values.filter { it % 2 == 1 }.joinToString(",")).toString(), "1,3,5")
            __check((values.dropWhile { it < 4 }.joinToString(",")).toString(), "4,5")
        }
