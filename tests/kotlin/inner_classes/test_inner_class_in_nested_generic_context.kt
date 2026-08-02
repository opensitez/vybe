// vybe-test: kotlin/inner_classes/test_inner_class_in_nested_generic_context
// origin: languages/kotlin/tests/kotlin/test_inner_classes.rs

class Generic<T>(val value: T) {
            inner class Holder {
                fun asString(): String = value.toString()
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Generic(12).Holder().asString()).toString(), "12")
        }
