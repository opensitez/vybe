// vybe-test: kotlin/member_references/test_reference_to_boolean_extension_function
// origin: languages/kotlin/tests/kotlin/test_member_references.rs

val check: (Boolean) -> String = Boolean::toString

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((check(true)).toString(), "true")
            __check((check(false)).toString(), "false")
        }
