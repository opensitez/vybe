// vybe-test: kotlin/data_class_copying/test_data_class_destructuring_copy_chain
// origin: languages/kotlin/tests/kotlin/test_data_class_copying.rs

data class Point(val x: Int, val y: Int)
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val p1 = Point(1, 2)
            val (x, y) = p1.copy(y = 10)
            __check((x).toString(), "1")
            __check((y).toString(), "10")
        }
