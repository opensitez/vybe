// vybe-test: kotlin/reified_generics/test_reified_type_check_for_int
// origin: languages/kotlin/tests/kotlin/test_reified_generics.rs

inline fun <reified T> isType(value: Any?): String {
            return if (value is T) "yes" else "no"
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((isType<Int>(3)).toString(), "yes")
            __check((isType<String>(3)).toString(), "no")
        }
