// vybe-test: kotlin/extension_functions/test_extension_receiver_can_be_nullable_and_differentiated
// origin: languages/kotlin/tests/kotlin/test_extension_functions.rs

fun String?.orEmptyTag(): String {
            return if (this == null) "none" else "ok:" + this
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val left: String? = null
            __check(("x".orEmptyTag()).toString(), "ok:x")
            __check((left.orEmptyTag()).toString(), "none")
        }
