// vybe-test: kotlin/enums/test_enum_when_nested_conditions
// origin: languages/kotlin/tests/kotlin/test_enums.rs

enum class Grade { A, B, C }
fun rank(g: Grade): Int { return when (g) { Grade.A -> 3
Grade.B -> 2
Grade.C -> 1 } }
fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { __check((rank(Grade.B)).toString(), "2")
__check((rank(Grade.A)).toString(), "3") }
