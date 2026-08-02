// vybe-test: kotlin/classes/test_class_with_custom_getter_and_state
// origin: languages/kotlin/tests/kotlin/test_classes.rs

class Box {
            val base: Int = 6
            val doubled: Int
                get() = base * 2
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val b = Box()
            __check((b.base).toString(), "6")
            __check((b.doubled).toString(), "12")
        }
