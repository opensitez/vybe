// vybe-test: kotlin/enums/test_enum_nested_when_without_default
// origin: languages/kotlin/tests/kotlin/test_enums.rs

enum class Tag { ONE, TWO }
fun describe(t: Tag): String { return when (t) { Tag.ONE -> "first"
Tag.TWO -> "second" } }
fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { __check((describe(Tag.ONE)).toString(), "first")
__check((describe(Tag.TWO)).toString(), "second") }
