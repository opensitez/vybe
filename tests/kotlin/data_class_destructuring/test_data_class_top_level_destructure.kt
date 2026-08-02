// vybe-test: kotlin/data_class_destructuring/test_data_class_top_level_destructure
// origin: languages/kotlin/tests/kotlin/test_data_class_destructuring.rs

data class PairVal(val x: Int, val y: Int)

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val (x, y) = PairVal(3, 7)
            __check((x).toString(), "3")
            __check((y).toString(), "7")
        }
