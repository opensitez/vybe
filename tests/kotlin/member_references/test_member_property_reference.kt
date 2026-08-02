// vybe-test: kotlin/member_references/test_member_property_reference
// origin: languages/kotlin/tests/kotlin/test_member_references.rs

class User(val id: Int)
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val readId = User::id
            __check((readId(User(9))).toString(), "9")
        }
