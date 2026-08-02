// vybe-test: kotlin/sealed_types/test_state_shape_preserved_in_when_mapping
// origin: languages/kotlin/tests/kotlin/test_sealed_types.rs

sealed class Variant {
            class A(val id: Int) : Variant()
            class B(val name: String) : Variant()
        }

        fun map(variant: Variant): String {
            return when (variant) {
                is Variant.A -> variant.id.toString()
                is Variant.B -> variant.name
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((map(Variant.A(4))).toString(), "4")
            __check((map(Variant.B("ok"))).toString(), "ok")
        }
