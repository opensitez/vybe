// vybe-test: kotlin/enums/test_enum_boolean_expression_match
// origin: languages/kotlin/tests/kotlin/test_enums.rs

enum class Flag { YES, NO }
fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val answer = Flag.YES
__check((if (answer == Flag.YES) "ok" else "no").toString(), "ok") }
