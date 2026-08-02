// vybe-test: kotlin/classes/test_class_data_like_projection_with_copy_behavior_not_available
// origin: languages/kotlin/tests/kotlin/test_classes.rs

class Pair(val left: Int, val right: Int)

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = Pair(1, 2)
            val b = Pair(a.left, a.right + 1)
            __check((a.left).toString(), "1")
            __check((a.right).toString(), "2")
            __check((b.right).toString(), "3")
        }
