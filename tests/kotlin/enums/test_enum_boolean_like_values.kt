// vybe-test: kotlin/enums/test_enum_boolean_like_values
// origin: languages/kotlin/tests/kotlin/test_enums.rs

enum class Switch { ON, OFF }
fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val s = Switch.OFF
__check((s == Switch.OFF).toString(), "true") }
