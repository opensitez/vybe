// vybe-test: kotlin/type_casts/test_is_check_on_generic_boxing_of_any
// origin: languages/kotlin/tests/kotlin/test_type_casts.rs

fun isStringList(value: Any): Boolean {
            return value is List<*>
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((isStringList(listOf("a", "b", "c"))).toString(), "true")
            __check((isStringList(10)).toString(), "false")
            val maybeList: Any? = null
            __check((maybeList is List<*>).toString(), "false")
        }
