// vybe-test: kotlin/generics/test_generic_type_bound_number
// origin: languages/kotlin/tests/kotlin/test_generics.rs

fun <T : Number> asInt(value: T): Int {
            return value.toInt()
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((asInt(12)).toString(), "12")
            __check((asInt(12.7)).toString(), "12")
        }
