// vybe-test: kotlin/generics/test_generic_array_projection_alias
// origin: languages/kotlin/tests/kotlin/test_generics.rs

fun <T> repeatThree(value: T): Array<T> {
            return arrayOf(value, value, value)
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = repeatThree("go")
            __check((values[0]).toString(), "go")
            __check((values[1]).toString(), "go")
            __check((values[2]).toString(), "go")
        }
