// vybe-test: kotlin/ranges/test_range_to_list_and_set_have_snapshot_semantics
// origin: languages/kotlin/tests/kotlin/test_ranges.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = (1..4).toList()
            val set = (1..4).toMutableList()
            set[0] = 9
            __check((values[0]).toString(), "1")
            __check((set[0]).toString(), "9")
            __check(((1..4).toMutableList()[0]).toString(), "1")
        }
