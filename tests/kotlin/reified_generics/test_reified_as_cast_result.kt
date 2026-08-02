// vybe-test: kotlin/reified_generics/test_reified_as_cast_result
// origin: languages/kotlin/tests/kotlin/test_reified_generics.rs

inline fun <reified T> asOrNull(value: Any?): String {
            val cast = value as? T
            return if (cast == null) "none" else "has"
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((asOrNull<String>("kotlin")).toString(), "has")
            __check((asOrNull<String>(8)).toString(), "none")
        }
