// vybe-test: kotlin/enums/test_enum_as_function_argument
// origin: languages/kotlin/tests/kotlin/test_enums.rs

enum class Mode { FAST, SLOW }
fun step(mode: Mode): Int { return if (mode == Mode.FAST) 2 else 1 }
fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { __check((step(Mode.FAST)).toString(), "2")
__check((step(Mode.SLOW)).toString(), "1") }
