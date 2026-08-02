// vybe-test: kotlin/enums/test_enum_set_membership_and_inclusion
// origin: languages/kotlin/tests/kotlin/test_enums.rs

enum class Flag { ON, OFF, UNKNOWN }
fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val active = Flag.ON
val allowed = setOf(Flag.ON, Flag.OFF)
__check((active in allowed).toString(), "true")
__check((Flag.UNKNOWN in allowed).toString(), "false") }
