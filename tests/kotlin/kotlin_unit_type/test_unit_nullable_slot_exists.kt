// vybe-test: kotlin/kotlin_unit_type/test_unit_nullable_slot_exists
// origin: languages/kotlin/tests/kotlin/test_kotlin_unit_type.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val x: Unit? = null
            __check((x == null).toString(), "true")
        }
