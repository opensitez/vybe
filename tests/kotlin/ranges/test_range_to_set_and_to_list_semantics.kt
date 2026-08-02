// vybe-test: kotlin/ranges/test_range_to_set_and_to_list_semantics
// origin: languages/kotlin/tests/kotlin/test_ranges.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val list = (1..4).toList()
            val set = (1..4).toSet()
            __check((list.size).toString(), "4")
            __check((list.joinToString()).toString(), "1, 2, 3, 4")
            __check((set.size).toString(), "4")
            __check((set.contains(4)).toString(), "true")
        }
