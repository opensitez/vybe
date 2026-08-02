// vybe-test: kotlin/extension_functions/test_extension_in_generic_bounded_receiver
// origin: languages/kotlin/tests/kotlin/test_extension_functions.rs

fun <T> T.describeIfString(default: String): String where T : Any? {
            return this?.toString() ?: default
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value: String? = null
            __check(("hello".describeIfString("none")).toString(), "hello")
            __check((value.describeIfString("none")).toString(), "none")
        }
