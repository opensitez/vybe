// vybe-test: kotlin/data_classes/test_data_class_copy_with_expression_args
// origin: languages/kotlin/tests/kotlin/test_data_classes.rs

data class Point(val x: Int, val y: Int, val z: Int = 0)

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = Point(1, 2)
            val b = a.copy(z = a.x + a.y)
            __check((b.z).toString(), "3")
            __check((b.x + b.y + b.z).toString(), "6")
        }
