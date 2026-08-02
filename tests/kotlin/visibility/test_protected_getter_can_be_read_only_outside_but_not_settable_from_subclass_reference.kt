// vybe-test: kotlin/visibility/test_protected_getter_can_be_read_only_outside_but_not_settable_from_subclass_reference
// origin: languages/kotlin/tests/kotlin/test_visibility.rs

open class Base {
            protected var value: Int = 1
        }

        class Child : Base() {
            var view: Int
                get() = value
                set(next) { value = next }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val child = Child()
            child.view = 10
            __check((child.view).toString(), "10")
        }
