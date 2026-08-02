// vybe-test: kotlin/class_delegation/test_class_delegation_with_custom_members
// origin: languages/kotlin/tests/kotlin/test_class_delegation.rs

interface Counter {
            fun value(): Int
        }

        class NumberCounter(private val v: Int) : Counter {
            override fun value() = v
        }

        class OffsetCounter(base: Counter) : Counter by base {
            fun id() = "offset"
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val c = OffsetCounter(NumberCounter(3))
            __check((c.id()).toString(), "offset")
            __check((c.value()).toString(), "3")
        }
