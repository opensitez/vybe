// vybe-test: kotlin/member_references/test_reference_through_lambda_map
// origin: languages/kotlin/tests/kotlin/test_member_references.rs

fun shout(x: Int) = x + 1
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val out = listOf(1, 2, 3).map(::shout).joinToString(";")
            __check((out).toString(), "2;3;4")
        }
