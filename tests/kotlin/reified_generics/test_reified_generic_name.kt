// vybe-test: kotlin/reified_generics/test_reified_generic_name
// origin: languages/kotlin/tests/kotlin/test_reified_generics.rs

inline fun <reified T> typeName(): String = T::class.simpleName ?: "unknown"

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((typeName<Int>()).toString(), "Int")
            __check((typeName<List<String>>()).toString(), "List")
        }
