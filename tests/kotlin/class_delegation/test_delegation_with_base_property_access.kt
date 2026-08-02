// vybe-test: kotlin/class_delegation/test_delegation_with_base_property_access
// origin: languages/kotlin/tests/kotlin/test_class_delegation.rs

interface Counter { val value: Int }

        class Base(override val value: Int) : Counter

        class Box(delegate: Counter) : Counter by delegate

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Box(Base(7)).value).toString(), "7")
        }
