// vybe-test: kotlin/data_classes/test_data_class_copy_modifies_single_field
// origin: languages/kotlin/tests/kotlin/test_data_classes.rs

data class Point(val x: Int, val y: Int)

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = Point(1, 2)
            val b = a.copy(y = 99)
            __check((b.x).toString(), "1")
            __check((b.y).toString(), "99")
        }
