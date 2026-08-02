// vybe-test: kotlin/generics/test_generic_writable_projection
// origin: languages/kotlin/tests/kotlin/test_generics.rs

fun <T> appendDefault(values: MutableList<in T>, value: T) {
            values.add(value)
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val numbers: MutableList<Number> = mutableListOf(1)
            appendDefault<Number>(numbers, 2)
            appendDefault<Number>(numbers, 3.5)
            __check((numbers[1]).toString(), "2")
            __check((numbers[2]).toString(), "3.5")
        }
