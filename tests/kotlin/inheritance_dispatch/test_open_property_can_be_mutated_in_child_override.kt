// vybe-test: kotlin/inheritance_dispatch/test_open_property_can_be_mutated_in_child_override
// origin: languages/kotlin/tests/kotlin/test_inheritance_dispatch.rs

open class Base {
            open var value: Int = 0
        }

        class Child : Base() {
            override var value: Int = 1
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val item = Child()
            item.value += 3
            __check((item.value).toString(), "4")
        }
