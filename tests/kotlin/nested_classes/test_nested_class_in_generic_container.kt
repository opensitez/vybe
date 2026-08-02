// vybe-test: kotlin/nested_classes/test_nested_class_in_generic_container
// origin: languages/kotlin/tests/kotlin/test_nested_classes.rs

class Bag<T> {
            class Entry
            class Typed<T>(val value: T)
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = Bag.Typed(2)
            val b = Bag.Entry()
            __check((a.value + 1).toString(), "3")
            __check((b::class.simpleName).toString(), "Entry")
        }
