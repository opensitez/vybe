// vybe-test: kotlin/member_references/test_standard_extension_property_reference
// origin: languages/kotlin/tests/kotlin/test_member_references.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val ref = String::length
            __check((ref("kotlin")).toString(), "6")
        }
