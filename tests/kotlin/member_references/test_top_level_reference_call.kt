// vybe-test: kotlin/member_references/test_top_level_reference_call
// origin: languages/kotlin/tests/kotlin/test_member_references.rs

fun square(v: Int) = v * v
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val f = ::square
            __check((f(7)).toString(), "49")
        }
