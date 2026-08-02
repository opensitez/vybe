// vybe-test: kotlin/data_classes/test_data_class_destructuring_with_nested_components_in_expression
// origin: languages/kotlin/tests/kotlin/test_data_classes.rs

data class Range(val start: Int, val end: Int)
        data class Window(val left: Range, val right: Range)

        fun size(window: Window): Int {
            val (first, second) = window
            return (first.end - first.start) + (second.end - second.start)
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val window = Window(Range(1, 4), Range(10, 20))
            __check((size(window)).toString(), "13")
        }
