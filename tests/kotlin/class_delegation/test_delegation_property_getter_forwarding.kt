// vybe-test: kotlin/class_delegation/test_delegation_property_getter_forwarding
// origin: languages/kotlin/tests/kotlin/test_class_delegation.rs

interface Counter {
            val count: Int
        }

        class State(override val count: Int) : Counter

        class Holder(delegate: Counter) : Counter by delegate

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = Holder(State(12))
            __check((value.count).toString(), "12")
        }
