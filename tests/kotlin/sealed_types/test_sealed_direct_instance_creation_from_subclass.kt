// vybe-test: kotlin/sealed_types/test_sealed_direct_instance_creation_from_subclass
// origin: languages/kotlin/tests/kotlin/test_sealed_types.rs

sealed class Unit {
            class Meter(val value: Int) : Unit()
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val unit = Unit.Meter(5)
            __check((unit.value).toString(), "5")
        }
