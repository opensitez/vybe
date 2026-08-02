// vybe-test: kotlin/member_references/test_top_level_reference_with_multiple_args
// origin: languages/kotlin/tests/kotlin/test_member_references.rs

fun add(a: Int, b: Int) = a + b
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val f = ::add
            __check((f(3, 4)).toString(), "7")
        }
