// vybe-test: kotlin/property_accessors/test_property_setter_and_getter_in_class_hierarchy
// origin: languages/kotlin/tests/kotlin/test_property_accessors.rs

open class Base {
            open var value: Int = 1
        }
        class Child : Base() {
            override var value: Int = 2
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val b: Base = Child()
            __check((b.value).toString(), "2")
        }
