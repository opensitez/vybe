// vybe-test: kotlin/member_references/test_reference_to_extension_receiver_function
// origin: languages/kotlin/tests/kotlin/test_member_references.rs

fun String.surround(left: String, right: String): String = left + this + right

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val ref: (String, String, String) -> String = String::surround
            __check((ref("k", "<", ">")).toString(), "<k>")
        }
