// vybe-test: kotlin/nested_classes/test_nested_class_in_object
// origin: languages/kotlin/tests/kotlin/test_nested_classes.rs

object Bridge {
            class Board(val size: Int)
            fun board(): Board = Board(4)
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Bridge.board().size).toString(), "4")
        }
