// vybe-test: kotlin/classes/test_class_method_overload_variation
// origin: languages/kotlin/tests/kotlin/test_classes.rs

class Math {
            fun value(x: Int): Int = x
            fun value(x: Int, y: Int): Int = x + y
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val m = Math()
            __check((m.value(1)).toString(), "1")
            __check((m.value(2, 3)).toString(), "5")
        }
