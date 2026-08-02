// vybe-test: kotlin/enums/test_enum_with_payload_chain_calc
// origin: languages/kotlin/tests/kotlin/test_enums.rs

enum class Level(val factor: Int) { LOW(1), MID(2), HIGH(3) }
fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val selected = Level.HIGH
__check((selected.factor * 2).toString(), "6") }
