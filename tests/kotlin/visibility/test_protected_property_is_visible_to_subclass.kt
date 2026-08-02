// vybe-test: kotlin/visibility/test_protected_property_is_visible_to_subclass
// origin: languages/kotlin/tests/kotlin/test_visibility.rs

open class Base {
            protected var value: Int = 1
        }

        class Child : Base() {
            fun bump() { value += 2 }
            fun read(): Int = value
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val child = Child()
            child.bump()
            __check((child.read()).toString(), "3")
        }
