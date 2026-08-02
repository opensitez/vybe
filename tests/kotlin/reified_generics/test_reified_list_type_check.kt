// vybe-test: kotlin/reified_generics/test_reified_list_type_check
// origin: languages/kotlin/tests/kotlin/test_reified_generics.rs

inline fun <reified T> hasStrings(values: List<Any>): String {
            return if (values.all { it is T }) "all" else "some"
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((hasStrings<String>(listOf("a", "b", "c"))).toString(), "all")
            __check((hasStrings<Int>(listOf("a", 1, "c"))).toString(), "some")
        }
