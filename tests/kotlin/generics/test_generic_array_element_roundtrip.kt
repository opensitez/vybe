// vybe-test: kotlin/generics/test_generic_array_element_roundtrip
// origin: languages/kotlin/tests/kotlin/test_generics.rs

fun <T> firstAndLast(values: Array<T>): String {
            return values.first().toString() + "," + values.last().toString()
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((firstAndLast(arrayOf(1, 2, 3))).toString(), "1,3")
            __check((firstAndLast(arrayOf("x", "y", "z"))).toString(), "x,z")
        }
