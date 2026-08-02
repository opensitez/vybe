// vybe-test: kotlin/member_references/test_reference_to_constructor_with_args
// origin: languages/kotlin/tests/kotlin/test_member_references.rs

class Pair(val left: Int, val right: String)

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val make = ::Pair
            val item = make(2, "x")
            __check((item.left).toString(), "2")
            __check((item.right).toString(), "x")
        }
