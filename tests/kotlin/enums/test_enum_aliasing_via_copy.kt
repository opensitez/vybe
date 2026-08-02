// vybe-test: kotlin/enums/test_enum_aliasing_via_copy
// origin: languages/kotlin/tests/kotlin/test_enums.rs

enum class Kind { A, B, C }
fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val a = Kind.A
val b = a
__check((a == b).toString(), "true") }
