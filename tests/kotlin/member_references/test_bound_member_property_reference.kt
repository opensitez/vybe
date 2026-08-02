// vybe-test: kotlin/member_references/test_bound_member_property_reference
// origin: languages/kotlin/tests/kotlin/test_member_references.rs

class User(val id: Int)
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val u = User(11)
            val read = u::id
            __check((read()).toString(), "11")
        }
