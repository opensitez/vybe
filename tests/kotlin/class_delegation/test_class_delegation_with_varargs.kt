// vybe-test: kotlin/class_delegation/test_class_delegation_with_varargs
// origin: languages/kotlin/tests/kotlin/test_class_delegation.rs

interface Summer {
            fun sum(values: IntArray): Int
        }

        class Adder : Summer {
            override fun sum(values: IntArray): Int = values.sum()
        }

        class SumWrapper(delegate: Summer) : Summer by delegate

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val wrapper = SumWrapper(Adder())
            __check((wrapper.sum(intArrayOf(1, 2, 3))).toString(), "6")
        }
