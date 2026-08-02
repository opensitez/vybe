// vybe-test: kotlin/classes/test_class_property_getter
// origin: languages/kotlin/tests/kotlin/test_classes.rs

class Box(val value: Int) {
            val doubled: Int
                get() = value * 2
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val box = Box(7)
            __check((box.doubled).toString(), "14")
        }
