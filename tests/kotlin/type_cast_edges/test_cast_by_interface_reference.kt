// vybe-test: kotlin/type_cast_edges/test_cast_by_interface_reference
// origin: languages/kotlin/tests/kotlin/test_type_cast_edges.rs

interface I { fun value(): Int }
        class Impl(val value: Int) : I { override fun value(): Int = value }

        fun asI(a: Any?): Int {
            return (a as? I)?.value() ?: -1
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((asI(Impl(7))).toString(), "7")
            __check((asI("x")).toString(), "-1")
        }
