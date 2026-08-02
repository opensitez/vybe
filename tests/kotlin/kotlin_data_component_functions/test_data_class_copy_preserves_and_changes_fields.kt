// vybe-test: kotlin/kotlin_data_component_functions/test_data_class_copy_preserves_and_changes_fields
// origin: languages/kotlin/tests/kotlin/test_kotlin_data_component_functions.rs

data class Point(val x: Int, val y: Int)

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val point = Point(2, 3)
            val shifted = point.copy(y = 9)
            __check((point).toString(), "Point(x=2, y=3)")
            __check((shifted).toString(), "Point(x=2, y=9)")
        }
