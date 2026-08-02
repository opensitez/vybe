// vybe-test: kotlin/generics/test_generic_list_first_or_default
// origin: languages/kotlin/tests/kotlin/test_generics.rs

fun <T> firstOrDefault(values: Array<T>, fallback: T): T {
            if (values.size == 0) {
                return fallback
            }
            return values[0]
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((firstOrDefault(arrayOf(4, 7, 9), 0)).toString(), "4")
            __check((firstOrDefault(arrayOf<String>(), "none")).toString(), "none")
        }
