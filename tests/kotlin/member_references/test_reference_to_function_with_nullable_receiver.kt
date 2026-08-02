// vybe-test: kotlin/member_references/test_reference_to_function_with_nullable_receiver
// origin: languages/kotlin/tests/kotlin/test_member_references.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val pick = String?.orEmpty
            __check((pick(null)).toString(), "")
            __check((pick("x")).toString(), "x")
        }
