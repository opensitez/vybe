// vybe-test: kotlin/inheritance_dispatch/test_getter_and_setter_overrides_are_used_from_base_reference
// origin: languages/kotlin/tests/kotlin/test_inheritance_dispatch.rs

open class Base {
            open var value: Int = 0
        }

        class Child : Base() {
            private var storage = 10

            override var value: Int
                get() = storage
                set(new_value) {
                    storage = new_value + 1
                }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val base: Base = Child()
            base.value = 7
            __check((base.value).toString(), "8")
            __check(((base as Child).value).toString(), "8")
        }
