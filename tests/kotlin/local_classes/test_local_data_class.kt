// vybe-test: kotlin/local_classes/test_local_data_class
// origin: languages/kotlin/tests/kotlin/test_local_classes.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            data class Point(val x: Int, val y: Int)
            val p = Point(1, 2)
            __check((p.x + p.y).toString(), "3")
        }
