// vybe-test: kotlin/data_classes/test_data_class_destructuring_by_index
// origin: languages/kotlin/tests/kotlin/test_data_classes.rs

data class PairValue(val left: Int, val right: Int)

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val p = PairValue(7, 11)
            val (left, right) = p
            __check((left + right).toString(), "18")
        }
