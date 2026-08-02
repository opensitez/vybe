// vybe-test: kotlin/functions/test_function_extension_member
// origin: languages/kotlin/tests/kotlin/test_functions.rs

fun String.wrap(prefix: String): String {
            return prefix + this
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check(("kotlin".wrap("[")).toString(), "[kotlin")
            __check(("v".wrap("v")).toString(), "vv")
        }
