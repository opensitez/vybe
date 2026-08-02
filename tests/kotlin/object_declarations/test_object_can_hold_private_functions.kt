// vybe-test: kotlin/object_declarations/test_object_can_hold_private_functions
// origin: languages/kotlin/tests/kotlin/test_object_declarations.rs

object Counter {
            private fun step(value: Int): Int = value + 1
            fun next(value: Int): Int = step(value)
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Counter.next(4)).toString(), "5")
        }
