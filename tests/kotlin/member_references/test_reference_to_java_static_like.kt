// vybe-test: kotlin/member_references/test_reference_to_java_static_like
// origin: languages/kotlin/tests/kotlin/test_member_references.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val fromInt = Int::toString
            __check((fromInt(5)).toString(), "5")
        }
