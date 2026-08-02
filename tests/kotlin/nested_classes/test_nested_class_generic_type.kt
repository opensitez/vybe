// vybe-test: kotlin/nested_classes/test_nested_class_generic_type
// origin: languages/kotlin/tests/kotlin/test_nested_classes.rs

class Holder<T> {
            class Slot<T>(val payload: T)
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val slot = Holder.Slot("text")
            __check((slot.payload).toString(), "text")
        }
