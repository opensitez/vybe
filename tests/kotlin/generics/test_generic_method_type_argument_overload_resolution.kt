// vybe-test: kotlin/generics/test_generic_method_type_argument_overload_resolution
// origin: languages/kotlin/tests/kotlin/test_generics.rs

fun <T> format(value: T): String = value.toString()

        fun <T> format(value: T, prefix: String): String {
            return prefix + ":" + value
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((format(9)).toString(), "9")
            __check((format(9, "num")).toString(), "num:9")
        }
