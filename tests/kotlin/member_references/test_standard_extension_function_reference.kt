// vybe-test: kotlin/member_references/test_standard_extension_function_reference
// origin: languages/kotlin/tests/kotlin/test_member_references.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val trimRef: (String) -> String = String::trim
            __check((trimRef("  a ")).toString(), "a")
        }
