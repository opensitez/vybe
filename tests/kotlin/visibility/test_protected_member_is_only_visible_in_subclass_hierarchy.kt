// vybe-test: kotlin/visibility/test_protected_member_is_only_visible_in_subclass_hierarchy
// origin: languages/kotlin/tests/kotlin/test_visibility.rs

open class Base {
            protected var value: Int = 3
        }

        class Child : Base() {
            fun write(next: Int) {
                value = next
            }

            fun read(): Int {
                return value
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val child = Child()
            child.write(6)
            __check((child.read()).toString(), "6")
        }
