// vybe-test: kotlin/data_classes/test_data_class_deeply_nested_destructuring
// origin: languages/kotlin/tests/kotlin/test_data_classes.rs

data class Line(val start: Int, val end: Int)
        data class Segment(val a: Line, val b: Line)

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val seg = Segment(Line(1, 2), Line(3, 4))
            val (left, right) = seg
            __check((left.start + right.end).toString(), "5")
            val (s, e) = left
            __check((s).toString(), "1")
            __check((e).toString(), "2")
        }
