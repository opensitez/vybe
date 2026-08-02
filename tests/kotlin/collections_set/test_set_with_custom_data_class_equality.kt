// vybe-test: kotlin/collections_set/test_set_with_custom_data_class_equality
// origin: languages/kotlin/tests/kotlin/test_collections_set.rs

data class PairKey(val left: Int, val right: Int)

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = setOf(PairKey(1, 2), PairKey(1, 2), PairKey(2, 1))
            __check((values.size).toString(), "2")
            __check((values.contains(PairKey(1, 2))).toString(), "true")
        }
