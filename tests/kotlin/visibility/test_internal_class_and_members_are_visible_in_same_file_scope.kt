// vybe-test: kotlin/visibility/test_internal_class_and_members_are_visible_in_same_file_scope
// origin: languages/kotlin/tests/kotlin/test_visibility.rs

internal class Vault(val value: Int)

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Vault(7).value).toString(), "7")
        }
