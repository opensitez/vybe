// vybe-test: kotlin/enums/test_enum_with_custom_payload_simple
// origin: languages/kotlin/tests/kotlin/test_enums.rs

enum class Size(val value: Int) { SMALL(1), MEDIUM(2), LARGE(3) }
fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { __check((Size.MEDIUM.value).toString(), "2")
__check((Size.LARGE.value).toString(), "3") }
