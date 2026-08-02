// vybe-test: kotlin/smart_casts/test_cast_to_common_supertype_then_refine_to_subtype
// origin: languages/kotlin/tests/kotlin/test_smart_casts.rs

open class Base
        class Left : Base()
        class Right : Base()

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val base: Base = Left()
            __check((base is Left).toString(), "true")
            __check((base is Right).toString(), "false")
            val value = base as? Left
            __check((value != null).toString(), "true")
        }
