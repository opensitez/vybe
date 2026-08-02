// vybe-test: kotlin/nested_classes/test_nested_class_basic_access
// origin: languages/kotlin/tests/kotlin/test_nested_classes.rs

class Box {
            class Inner(val value: Int)
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val item = Box.Inner(7)
            __check((item.value).toString(), "7")
        }
