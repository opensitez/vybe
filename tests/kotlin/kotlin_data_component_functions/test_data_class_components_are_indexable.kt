// vybe-test: kotlin/kotlin_data_component_functions/test_data_class_components_are_indexable
// origin: languages/kotlin/tests/kotlin/test_kotlin_data_component_functions.rs

data class Point(val x: Int, val y: Int)

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val point = Point(4, 7)
            __check((point.component1()).toString(), "4")
            __check((point.component2()).toString(), "7")
        }
